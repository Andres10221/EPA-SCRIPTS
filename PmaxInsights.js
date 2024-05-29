{\rtf1\ansi\ansicpg1252\cocoartf2709
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 ArialMT;\f1\froman\fcharset0 Times-Roman;}
{\colortbl;\red255\green255\blue255;\red0\green0\blue0;}
{\*\expandedcolortbl;;\cssrgb\c0\c0\c0;}
\margl1440\margr1440\vieww28600\viewh14980\viewkind0
\deftab720
\pard\pardeftab720\partightenfactor0

\f0\fs21\fsmilli10667 \cf0 \expnd0\expndtw0\kerning0
\outl0\strokewidth0 \strokec2 // v56 - added title data aggregated at account level (not just campaign level)
\f1\fs24 \

\f0\fs21\fsmilli10667 // added productTitles for Whisperer script
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // template sheet to copy -> https://docs.google.com/spreadsheets/d/1ibBae_QEa_o4R2Jtld_qyApimDuvYzx3y2TiIh8wOus/copy
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // user config section - make your changes here
\f1\fs24 \
\
\

\f0\fs21\fsmilli10667 const SHEET_URL\'a0 \'a0 = 'https://docs.google.com/spreadsheets/d/1qRx06DR5xh0KV5DqYICzD9iVxNhDPV1R85-EpJ2nAus/edit#gid=1441562279'\'a0 \'a0 // create a copy of the template above first
\f1\fs24 \

\f0\fs21\fsmilli10667 const clientCode \'a0 = 'COP'\'a0 \'a0 // this string will be added to the sheet name\'a0
\f1\fs24 \
\
\
\

\f0\fs21\fsmilli10667 // for all these variables, if you enter a value on the 'Settings' tab in the sheet, it will override the default values here
\f1\fs24 \

\f0\fs21\fsmilli10667 let defaultNumberOfDays = 30\'a0 \'a0 \'a0 // how many days data do you want to get?\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 let defaulttCost\'a0 \'a0 \'a0 \'a0 = 10\'a0 \'a0 \'a0 // the 'threshold' for low vs high cost (as a proxy for importance)
\f1\fs24 \

\f0\fs21\fsmilli10667 let defaulttRoas\'a0 \'a0 \'a0 \'a0 = 4 \'a0 \'a0 \'a0 // the 'threshold' for low vs high roas (as a proxy for performance)
\f1\fs24 \

\f0\fs21\fsmilli10667 let defaultBrandTerm\'a0 \'a0 = ''\'a0 \'a0 \'a0 // a string that best represents the brand name - used to find the search category 'buckets' - use a comma separated list for multiple brand terms
\f1\fs24 \
\
\

\f0\fs21\fsmilli10667 // wiki with instructions: https://mikerhodes.notion.site/mikerhodes/83cbff0126764ac7aaed40bdb37c8a00?v=6476728e4a78439badc0f63e961d52bf
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // please don\'92t change any code below this line \'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97\'97
\f1\fs24 \

\f0\fs21\fsmilli10667 // no really
\f1\fs24 \
\

\f0\fs21\fsmilli10667 function main() \{
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let startDuration = new Date();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Setup Sheet
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0Logger.log(`Starting the PMax Insights script. Building queries...`);
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let accountName = AdsApp.currentAccount().getName();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let ident = clientCode ? clientCode : accountName;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let ss = SpreadsheetApp.openByUrl(SHEET_URL);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0ss.rename(ident + ' - v56 PMax Insights - MikeRhodes.com.au (c) ');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let defaultSettings = \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0numberOfDays: defaultNumberOfDays,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0tCost: defaulttCost,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0tRoas: defaulttRoas,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0brandTerm: defaultBrandTerm,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Update the default settings with the values from the sheet if available
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let settings = updateVariablesFromSheet(ss, defaultSettings);
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let numberOfDays= settings.numberOfDays;\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let tCost \'a0 \'a0 \'a0 = settings.tCost;\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let tRoas \'a0 \'a0 \'a0 = settings.tRoas;\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let brandTerm \'a0 = settings.brandTerm;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let fromDate\'a0 \'a0 = settings.fromDate;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let toDate\'a0 \'a0 \'a0 = settings.toDate;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let lotsProducts= settings.lotsProducts;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let turnonTitle = settings.turnonTitle;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let turnonID\'a0 \'a0 = settings.turnonID;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let turnonTitleAccount = settings.turnonTitleAccount;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let campFilter\'a0 = settings.campFilter;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let turnonPlace = settings.turnonPlace;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let turnonChange= settings.turnonChange;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let turnonOpenai= settings.turnonOpenai;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let productDays = settings.productDays;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let levenshtein = 2;\'a0
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let settingsWithoutApiKey = \{...settings\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0delete settingsWithoutApiKey.apiKey;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0Logger.log('Settings updated: ' + JSON.stringify(settingsWithoutApiKey));
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// define query elements. wrap with spaces for safety
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let impr \'a0 \'a0 \'a0 = ' metrics.impressions ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let clicks \'a0 \'a0 = ' metrics.clicks ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let cost \'a0 \'a0 \'a0 = ' metrics.cost_micros ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let engage \'a0 \'a0 = ' metrics.engagements ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let inter\'a0 \'a0 \'a0 = ' metrics.interactions ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let conv \'a0 \'a0 \'a0 = ' metrics.conversions ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let value\'a0 \'a0 \'a0 = ' metrics.conversions_value ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let allConv\'a0 \'a0 = ' metrics.all_conversions ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let allValue \'a0 = ' metrics.all_conversions_value ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let views\'a0 \'a0 \'a0 = ' metrics.video_views ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let cpv\'a0 \'a0 \'a0 \'a0 = ' metrics.average_cpv ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let eventTypes = ' metrics.interaction_event_types ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let segDate\'a0 \'a0 = ' segments.date ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let prodTitle\'a0 = ' segments.product_title ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let prodID \'a0 \'a0 = ' segments.product_item_id ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let campName \'a0 = ' campaign.name ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let campId \'a0 \'a0 = ' campaign.id ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let chType \'a0 \'a0 = ' campaign.advertising_channel_type ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let aIdAsset \'a0 = ' asset.resource_name ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let aId\'a0 \'a0 \'a0 \'a0 = ' asset.id ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let assetSource= ' asset.source ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let agId \'a0 \'a0 \'a0 = ' asset_group.id ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let assetFtype = ' asset_group_asset.field_type ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let adPmaxPerf = ' asset_group_asset.performance_label ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let agStrength = ' asset_group.ad_strength ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let agStatus \'a0 = ' asset_group.status ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let agPrimary\'a0 = ' asset_group.primary_status ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let asgName\'a0 \'a0 = ' asset_group.name ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let lgType \'a0 \'a0 = ' asset_group_listing_group_filter.type ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let aIdCamp\'a0 \'a0 = ' segments.asset_interaction_target.asset ';\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let assetName\'a0 = ' asset.name ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let adUrl\'a0 \'a0 \'a0 = ' asset.image_asset.full_size.url ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let ytTitle\'a0 \'a0 = ' asset.youtube_video_asset.youtube_video_title ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let ytId \'a0 \'a0 \'a0 = ' asset.youtube_video_asset.youtube_video_id ';\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let placement\'a0 = ' detail_placement_view.group_placement_target_url ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let placeType\'a0 = ' detail_placement_view.placement_type ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let chgDateTime= ' change_event.change_date_time ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let chgResType = ' change_event.change_resource_type ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let chgFields\'a0 = ' change_event.changed_fields ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let clientTyp\'a0 = ' change_event.client_type ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let feed \'a0 \'a0 \'a0 = ' change_event.feed ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let feedItm\'a0 \'a0 = ' change_event.feed_item ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let newRes \'a0 \'a0 = ' change_event.new_resource ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let oldRes \'a0 \'a0 = ' change_event.old_resource ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let resChgOp \'a0 = ' change_event.resource_change_operation ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let resName\'a0 \'a0 = ' change_event.resource_name ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let userEmail\'a0 = ' change_event.user_email ';\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let interAsset = ' segments.asset_interaction_target.interaction_on_this_asset ';\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let pMaxOnly \'a0 = ' AND campaign.advertising_channel_type = "PERFORMANCE_MAX" ';\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let pMaxShop \'a0 = ' AND campaign.advertising_channel_type IN ("SHOPPING","PERFORMANCE_MAX") ';\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let campLike \'a0 = ` AND campaign.name LIKE "%$\{campFilter\}%" `;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let agFilter \'a0 = ' AND asset_group_listing_group_filter.type != "SUBDIVISION" ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let notInter \'a0 = ' AND segments.asset_interaction_target.interaction_on_this_asset != "TRUE" ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let order\'a0 \'a0 \'a0 = ' ORDER BY campaign.name ';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Date stuff --------------------------------------------
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Get the time zone of the current Google Ads account
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let timeZone = AdsApp.currentAccount().getTimeZone();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Check if fromDate and toDate are defined before using them
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let dateSwitch = fromDate !== undefined && toDate !== undefined ? 1 : 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let today = new Date(), yesterday = new Date(), startDate = new Date();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0yesterday.setDate(today.getDate() - 1);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0startDate.setDate(today.getDate() - numberOfDays);\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// format the dates to be used in the queries
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let formattedYesterday = Utilities.formatDate(yesterday, timeZone, 'yyyy-MM-dd');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let formattedStartDate = Utilities.formatDate(startDate, timeZone, 'yyyy-MM-dd');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Function to ensure the date from the sheet is treated as a literal string
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0function formatDateLiteral(dateString) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// Use a regular expression to extract date parts
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let dateParts = dateString.match(/(\\d\{2\})\\/(\\d\{2\})\\/(\\d\{4\})/);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (!dateParts) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0throw new Error('Date format is not valid. Expected format dd/mm/yyyy.');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// Rearrange the date parts to 'yyyy-MM-dd' format
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let formattedDate = `$\{dateParts[3]\}-$\{dateParts[2]\}-$\{dateParts[1]\}`;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0return formattedDate;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Use the function to format fromDate and toDate as literal strings
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let formattedFromDate = dateSwitch ? formatDateLiteral(fromDate) : undefined;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let formattedToDate \'a0 = dateSwitch ? formatDateLiteral(toDate) : undefined;\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// set range for the queries - if switch is 1 use the from/to dates
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let mainDateRange = dateSwitch ? `segments.date BETWEEN "$\{formattedFromDate\}" AND "$\{formattedToDate\}"` : `segments.date BETWEEN "$\{formattedStartDate\}" AND "$\{formattedYesterday\}"`;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Build queries --------------------------------------------
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let assetGroupAssetColumns = [campName, asgName, agId, aIdAsset, assetFtype, campId]; //\'a0 adPmaxPerf, agStrength, agStatus, assetSource,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let assetGroupAssetQuery = 'SELECT ' + assetGroupAssetColumns.join(',') +
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0' FROM asset_group_asset ' +\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0' WHERE asset_group_asset.field_type NOT IN ("HEADLINE", "DESCRIPTION", "LOGO", "LANDSCAPE_LOGO", "BUSINESS_NAME", "LONG_HEADLINE", "CALL_TO_ACTION_SELECTION")' + pMaxShop + campLike;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let displayVideoColumns = [segDate, campName, aIdCamp, cost, conv, value, views, cpv, impr, clicks, chType, interAsset, campId];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let displayVideoQuery = 'SELECT ' + displayVideoColumns.join(',') +
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0' FROM campaign ' +
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0' WHERE ' + mainDateRange + pMaxShop + campLike + notInter + order;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let assetGroupColumns = [segDate, campName, asgName, agStrength, agStatus, lgType, impr, clicks, cost, conv, value];\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let assetGroupQuery = 'SELECT ' + assetGroupColumns.join(',')\'a0 +\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0' FROM asset_group_product_group_view ' +
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0' WHERE ' + mainDateRange + agFilter + campLike ;
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let campaignColumns = [segDate, campName, cost, conv, value, views, cpv, impr, clicks, chType, campId];\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let campaignQuery = 'SELECT ' + campaignColumns.join(',') +\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0' FROM campaign ' +
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0' WHERE ' + mainDateRange + pMaxShop + campLike + order;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let assetColumns = [aIdAsset, assetSource, ytTitle, ytId, assetName]\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let assetQuery = 'SELECT ' + assetColumns.join(',')\'a0 +\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0' FROM asset ' +\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0' WHERE asset.type IN ("YOUTUBE_VIDEO", "IMAGE")'
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let assetGroupMetrics = [campName, asgName, cost, conv, value, impr, clicks, agPrimary];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let assetGroupMetricsQuery = 'SELECT ' + assetGroupMetrics.join(',')\'a0 +
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0' FROM asset_group ' +
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0' WHERE ' + mainDateRange + ' AND metrics.impressions > 0 ';
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let changeColumns = [campName, userEmail, chgDateTime, chgResType, chgFields, clientTyp, feed, feedItm, newRes, oldRes, resChgOp];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let changeQuery = 'SELECT ' + changeColumns.join(',') +
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0' FROM change_event ' +
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0' WHERE change_event.change_date_time DURING LAST_7_DAYS ' + pMaxShop + campLike +
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0' ORDER BY change_event.change_date_time DESC ' +
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0' LIMIT 9999';\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let placeColumns = [campName, placement, placeType, impr, inter, views, cost, conv, value, allConv, allValue]
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let placeQuery = 'SELECT ' + placeColumns.join(',') +
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0' FROM detail_placement_view ' +
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0' WHERE ' + mainDateRange + campLike + order;
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let productColumns = [prodTitle, prodID, cost, conv, value, impr, clicks, campName, chType];\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let productQuery = 'SELECT ' + productColumns.join(',')\'a0 +\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0' FROM shopping_performance_view\'a0 ' +\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0' WHERE metrics.impressions > 0 AND ' + mainDateRange + pMaxShop + campLike ;\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// SINGLE account snippet
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// modify the productQuery based on the value of lotsProducts
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0switch (lotsProducts) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0case 1:
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0productQuery += ' AND metrics.cost_micros > 0';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0break;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0case 2:
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0productQuery += ' AND metrics.conversions > 0';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0break;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Process data --------------------------------------------
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let campaignData = fetchData(campaignQuery);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// If campaignData is empty, no eligible PMax/Shopping campaigns were found
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0if (campaignData.length === 0) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0Logger.log('No eligible PMax campaigns found.');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0return; // Exit the script if no eligible campaigns were found
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0Logger.log('Fetching campaign data. This may take a few moments...');
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let titleData \'a0 \'a0 \'a0 \'a0 \'a0 = fetchProductData(productQuery, tCost, tRoas, 'prodTitle');\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let idData\'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 = fetchProductData(productQuery, tCost, tRoas, 'idChannel');\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let assetGroupAssetData = fetchData(assetGroupAssetQuery);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let displayVideoData\'a0 \'a0 = fetchData(displayVideoQuery);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let assetGroupData\'a0 \'a0 \'a0 = fetchData(assetGroupQuery);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let assetData \'a0 \'a0 \'a0 \'a0 \'a0 = fetchData(assetQuery);\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let assetGroupNewData \'a0 = fetchData(assetGroupMetricsQuery);
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0Logger.log('Data fetching complete.\'a0 Now processing data...');
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Extract marketing assets & de-dupe
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let \{ displayAssets, videoAssets \} = extractAndDeduplicateAssets(assetGroupAssetData);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Filter displayVideoData based on the type
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let displayAssetData = filterDataByAssets(displayVideoData, displayAssets);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let videoAssetData \'a0 = filterDataByAssets(displayVideoData, videoAssets);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Process data for each dataset
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let processedDisplayAssetData = aggregateDataByDateAndCampaign(displayAssetData);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let processedVideoAssetData \'a0 = aggregateDataByDateAndCampaign(videoAssetData);\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let processedAssetGroupData \'a0 = processData(assetGroupData);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let processedCampData \'a0 \'a0 \'a0 \'a0 = processData(campaignData);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Combine all non-search metrics, calc 'search' & process summary
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let nonSearchData = [...processedDisplayAssetData, ...processedVideoAssetData, ...processedAssetGroupData];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let searchResults = getSearchResults(processedCampData, nonSearchData);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let totalData \'a0 \'a0 = processTotalData(processedCampData, processedAssetGroupData, processedDisplayAssetData, processedVideoAssetData, searchResults);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let summaryData \'a0 = processSummaryData(processedCampData, processedAssetGroupData, processedDisplayAssetData, processedVideoAssetData, searchResults);\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Aggregate the metrics for display and video assets
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let aggregatedDisplayAssetMetrics = aggregateMetricsByAsset(displayAssetData);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let aggregatedVideoAssetMetrics \'a0 = aggregateMetricsByAsset(videoAssetData);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let enrichedDisplayAssetDetails \'a0 = enrichAssetMetrics(aggregatedDisplayAssetMetrics, assetData, 'display');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let enrichedVideoAssetDetails \'a0 \'a0 = enrichAssetMetrics(aggregatedVideoAssetMetrics, assetData, 'video');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let mergedDisplayData \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 = mergeMetricsWithDetails(aggregatedDisplayAssetMetrics, enrichedDisplayAssetDetails);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let mergedVideoData \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 = mergeMetricsWithDetails(aggregatedVideoAssetMetrics, enrichedVideoAssetDetails);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Output the data to respective sheets
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0Logger.log('Data processing complete. Now writing data to the sheets...');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0outputAggregatedDataToSheet(ss, 'display', \'a0 \'a0 mergedDisplayData);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0outputAggregatedDataToSheet(ss, 'video', \'a0 \'a0 \'a0 mergedVideoData);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0outputSummaryToSheet( \'a0 \'a0 \'a0 ss, 'totals',\'a0 \'a0 \'a0 totalData);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0outputSummaryToSheet( \'a0 \'a0 \'a0 ss, 'summary', \'a0 \'a0 summaryData);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0outputDataToSheet(\'a0 \'a0 \'a0 \'a0 \'a0 ss, 'group', \'a0 \'a0 \'a0 assetGroupData);\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0outputDataToSheet(\'a0 \'a0 \'a0 \'a0 \'a0 ss, 'prodTitle', \'a0 titleData);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0outputDataToSheet(\'a0 \'a0 \'a0 \'a0 \'a0 ss, 'idChannel', \'a0 idData);\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0outputDataToSheet(\'a0 \'a0 \'a0 \'a0 \'a0 ss, 'assetGroup',\'a0 assetGroupNewData);\'a0 \'a0 \'a0 // asset group data with date range (v54)
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0extractTerms( \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 ss, mainDateRange, brandTerm, levenshtein); // get search categories\'a0\'a0
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// not used yet ------------------
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0//extractNGrams(ss, mainDateRange);\'a0
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// advanecd settings flags - use the hidden 'Advanced' & 'Super Advanced' tabs to control these
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let titleIDData, idAccData, titleAccountData, placeData, changeData;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0if (turnonTitle) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let titleIDData = fetchProductData(productQuery, tCost, tRoas, 'prodTitleID');\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0outputDataToSheet(ss, 'prodTitleID', titleIDData);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0if (turnonID) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let idAccData = fetchProductData(productQuery, tCost, tRoas, 'idAccount');\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0outputDataToSheet(ss, 'idAccount', idAccData);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0if (turnonTitleAccount) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let titleAccountData = fetchProductData(productQuery, tCost, tRoas, 'prodAccount');\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0outputDataToSheet(ss, 'prodAccount', titleAccountData);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0if (turnonPlace) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let placeReport = AdsApp.report(placeQuery);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let rows = placeReport.rows();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let aggData = aggMet(rows);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0outputPlaceToSheet(ss, aggData, numberOfDays);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0if (turnonChange) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let changeData = fetchData(changeQuery);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0outputDataToSheet(ss, 'changeData', changeData);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0if (turnonOpenai) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0createTitleReport(ss, 'Report', settings);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let endDuration\'a0 = new Date();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let duration = (endDuration - startDuration) / 1000;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0log(ss, duration, settings, ident);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0Logger.log(`Script execution time: $\{duration\} seconds. Finished script for $\{ident\}.`);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// MCC snippet
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Check if campaignData is empty, which means no eligible PMax campaigns were found
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0/* if (typeof totalData === 'undefined' || totalData.length === 0) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0Logger.log('No eligible PMax/Shopping campaigns found.');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0return; // Exit the script if no eligible PMax campaigns are found
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\} else \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0clientSettings.totalData = totalData; // pass totals back up the chain
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0return clientSettings;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\} */
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \} // end main func
\f1\fs24 \
\
\
\

\f0\fs21\fsmilli10667 // various additional functions ---------------------------------------------
\f1\fs24 \
\
\

\f0\fs21\fsmilli10667 // function to create LLM report on product titles, calling openai API
\f1\fs24 \

\f0\fs21\fsmilli10667 function createTitleReport(ss, tabName, s) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0Logger.log('Generating report with OpenAI');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let productTitlesSheet = ss.getSheetByName('productTitles');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// prompt is in a named range called p_prompt, data is in a named range called p_data
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let promptRange = ss.getRangeByName('p_product');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let profitableRange = ss.getRangeByName('p_profitable');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let costlyRange = ss.getRangeByName('p_costly');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let prompt, profitable, costly;
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0if (promptRange && profitableRange && costlyRange) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0prompt = promptRange.getValues().flat().filter(String).join("\\n");
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0profitable = profitableRange.getValues().flat().filter(String).join("\\n");
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0costly = costlyRange.getValues().flat().filter(String).join("\\n");
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\} else \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0Logger.log('Error: Prompt or Profitable/Costly ranges are empty. Try using a longer timeframe to get more product data');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0return;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let url = 'https://api.openai.com/v1/chat/completions';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// choose openai model to use based on s.model (better or cheaper)
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let model = s.model == 'better' ? 'gpt-4-turbo-preview' : 'gpt-3.5-turbo'
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let messages = [
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\{ "role": "user", "content": prompt + " " + profitable + " " + costly\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0];
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let payload = \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0"model": model,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0"messages": messages
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let httpOptions = \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0"method": "POST",
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0"muteHttpExceptions": true,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0"contentType": "application/json",
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0"headers": \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0"Authorization": 'Bearer ' + s.apiKey
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\},
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'payload': JSON.stringify(payload)
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let response = UrlFetchApp.fetch(url, httpOptions);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let startTime = Date.now();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0while (response.getResponseCode() !== 200 && Date.now() - startTime < 60000) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0Utilities.sleep(10000);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0response = UrlFetchApp.fetch(url, httpOptions);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0Logger.log('Time elapsed: ' + (Date.now() - startTime) / 1000 + ' seconds');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0if (response.getResponseCode() !== 200) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0Logger.log('Error: OpenAI API request timed out after 60 seconds. Please try again.');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0return;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let responseJson = JSON.parse(response.getContentText());
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let choices = responseJson.choices;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let text = choices[0].message.content;
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0Logger.log(text)
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// output text to named range 'openaiOutput' in tabName (Report)
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let openaiOutput = ss.getRangeByName('openaiOutput');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0openaiOutput.setValue(text);
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // flatten the object data to enable more processing of that data
\f1\fs24 \

\f0\fs21\fsmilli10667 function flattenObject(ob) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let toReturn = \{\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0for (let i in ob) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if ((typeof ob[i]) === 'object') \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let flatObject = flattenObject(ob[i]);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0for (let x in flatObject) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0toReturn[i + '.' + x] = flatObject[x];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\} else \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0toReturn[i] = ob[i];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0return toReturn;
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // fetch data given a query string using search (not report)
\f1\fs24 \

\f0\fs21\fsmilli10667 function fetchData(q) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let data = [];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0iterator = AdsApp.search(q, \{'apiVersion': 'v16'\});
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0while (iterator.hasNext()) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const row = iterator.next();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const rowData = flattenObject(row); // Flatten the row data
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0data.push(rowData);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0return data;
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // extract & dedupe in one step
\f1\fs24 \

\f0\fs21\fsmilli10667 function extractAndDeduplicateAssets(data) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let displayAssetsSet = new Set();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let videoAssetsSet = new Set();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0data.forEach(row => \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (row['assetGroupAsset.fieldType'] && row['assetGroupAsset.fieldType'].includes('MARKETING')) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0displayAssetsSet.add(row['asset.resourceName']);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (row['assetGroupAsset.fieldType'] && row['assetGroupAsset.fieldType'].includes('VIDEO')) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0videoAssetsSet.add(row['asset.resourceName']);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\});
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0return \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0displayAssets: [...displayAssetsSet],
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0videoAssets: [...videoAssetsSet]
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // filter the display & video assets
\f1\fs24 \

\f0\fs21\fsmilli10667 function filterDataByAssets(data, assets) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0return data.filter(row => assets.includes(row['segments.assetInteractionTarget.asset']));
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // get additional data from asset for video\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 function mergeMetricsWithDetails(aggregatedVideoAssetMetrics, enrichedVideoAssetDetails) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0return enrichedVideoAssetDetails.map(detail => \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const metrics = aggregatedVideoAssetMetrics[detail.assetName];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0return \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0...detail,\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0...metrics
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\});
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // enrich assets
\f1\fs24 \

\f0\fs21\fsmilli10667 function enrichAssetMetrics(aggregatedMetrics, assetData, type) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let assetDetailsArray = [];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// For each asset in aggregatedMetrics, fetch details from assetData
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0for (let assetName of Object.keys(aggregatedMetrics)) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// Find the asset in assetData
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let matchingAsset = assetData.find(asset => asset['asset.resourceName'] === assetName);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (matchingAsset) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let assetDetails = \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0assetName: assetName,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0assetSource: matchingAsset['asset.source'],
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0assetImage: matchingAsset['asset.name']\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (type === 'video') \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0assetDetails.youtubeTitle = matchingAsset['asset.youtubeVideoAsset.youtubeVideoTitle'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0assetDetails.youtubeId = matchingAsset['asset.youtubeVideoAsset.youtubeVideoId'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0assetDetailsArray.push(assetDetails);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0return assetDetailsArray;
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // Aggregate display & video assets to show total metrics for each asset in their sheets
\f1\fs24 \

\f0\fs21\fsmilli10667 function aggregateMetricsByAsset(data) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0const aggregatedData = \{\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0data.forEach(row => \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const asset = row['segments.assetInteractionTarget.asset'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (!aggregatedData[asset]) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregatedData[asset] = \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0clicks: 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0videoViews: 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0conversionsValue: 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0conversions: 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0cost: 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0impressions: 0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregatedData[asset].clicks += parseInt(row['metrics.clicks']);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregatedData[asset].videoViews += parseInt(row['metrics.videoViews']);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregatedData[asset].conversionsValue += parseFloat(row['metrics.conversionsValue']);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregatedData[asset].conversions += parseInt(row['metrics.conversions']);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregatedData[asset].cost += parseInt(row['metrics.costMicros']) / 1e6;\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregatedData[asset].impressions += parseInt(row['metrics.impressions']);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\});
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0return aggregatedData;
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // function to output the display & video metrics to tabs
\f1\fs24 \

\f0\fs21\fsmilli10667 function outputAggregatedDataToSheet(ss, sheetName, data) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Create or access the sheet
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Clear the sheet
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0sheet.clear();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Set the header based on sheetName
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0const headers = sheetName === 'video' ?
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0[['Asset Name', 'Source', 'YouTube Title', 'YouTube ID', 'Impr', 'Clicks', 'Views', 'Value', 'Conv', 'Cost']] :
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0[['Asset Name', 'Source', 'ImageName', 'Impr', 'Clicks', 'Views', 'Value', 'Conv', 'Cost']];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Build the data array
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0const dataArray = data.map(item => \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0return sheetName === 'video' ?
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0[item.assetName, item.assetSource, item.youtubeTitle, item.youtubeId, item.impressions, item.clicks, item.videoViews, item.conversionsValue, item.conversions, item.cost] :
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0[item.assetName, item.assetSource, item.assetImage, item.impressions, item.clicks, item.videoViews, item.conversionsValue, item.conversions, item.cost];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\});
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Combine headers and data
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0const fullData = headers.concat(dataArray);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Set the values to the sheet
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0sheet.getRange(1, 1, fullData.length, fullData[0].length).setValues(fullData);
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // Process data: Aggregate by date, campaign name, and various metrics
\f1\fs24 \

\f0\fs21\fsmilli10667 function aggregateDataByDateAndCampaign(data) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0const aggregatedData = \{\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0data.forEach(row => \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const date = row['segments.date'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const campaignName = row['campaign.name'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const key = `$\{date\}_$\{campaignName\}`;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (!aggregatedData[key]) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregatedData[key] = \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'date': date,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'campaignName': campaignName,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'cost': 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'impressions': 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'clicks': 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'conversions': 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'conversionsValue': 0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (row['metrics.costMicros']) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregatedData[key].cost += row['metrics.costMicros'] / 1000000;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (row['metrics.impressions']) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregatedData[key].impressions += parseInt(row['metrics.impressions']);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (row['metrics.clicks']) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregatedData[key].clicks += parseInt(row['metrics.clicks']);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (row['metrics.conversions']) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregatedData[key].conversions += parseFloat(row['metrics.conversions']);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (row['metrics.conversionsValue']) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregatedData[key].conversionsValue += parseFloat(row['metrics.conversionsValue']);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\});
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0return Object.values(aggregatedData);
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // calc the search results
\f1\fs24 \

\f0\fs21\fsmilli10667 function getSearchResults(processedCampData, nonSearchData) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Initialize an object to hold 'search' metrics
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0const searchMetrics = \{\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Sum up the non-search metrics
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0nonSearchData.forEach(row => \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const key = `$\{row.date\}_$\{row.campaignName\}`;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (!searchMetrics[key]) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0searchMetrics[key] = \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0date: row.date,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0campaignName: row.campaignName,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0cost: 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0impressions: 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0clicks: 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0conversions: 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0conversionsValue: 0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0for (let metric of ['cost', 'impressions', 'clicks', 'conversions', 'conversionsValue']) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0searchMetrics[key][metric] += row[metric];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\});
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Calculate 'search' metrics
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0const searchResults = [];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0processedCampData.forEach(row => \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const rowCopy = JSON.parse(JSON.stringify(row));\'a0 // Deep copy
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const key = `$\{row.date\}_$\{row.campaignName\}`;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (searchMetrics[key]) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0for (let metric of ['cost', 'impressions', 'clicks', 'conversions', 'conversionsValue']) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0rowCopy[metric] -= (searchMetrics[key][metric] || 0);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// replace negative values with 0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (rowCopy[metric] < 0) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0rowCopy[metric] = 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0searchResults.push(rowCopy);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\});
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0return searchResults;
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // process the data
\f1\fs24 \

\f0\fs21\fsmilli10667 function processData(data) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0const summedData = \{\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0data.forEach(row => \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const date = row['segments.date'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const campaignName = row['campaign.name'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const campaignType = row['campaign.advertisingChannelType'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const key = `$\{date\}_$\{campaignName\}`;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// Initialize if the key doesn't exist
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (!summedData[key]) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0summedData[key] = \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'date': date,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'campaignName': campaignName,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'campaignType': campaignType,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'cost': 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'impressions': 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'clicks': 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'conversions': 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'conversionsValue': 0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (row['metrics.costMicros']) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0summedData[key].cost += row['metrics.costMicros'] / 1000000;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (row['metrics.impressions']) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0summedData[key].impressions += parseInt(row['metrics.impressions']);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (row['metrics.clicks']) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0summedData[key].clicks += parseInt(row['metrics.clicks']);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (row['metrics.conversions']) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0summedData[key].conversions += parseFloat(row['metrics.conversions']);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (row['metrics.conversionsValue']) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0summedData[key].conversionsValue += parseFloat(row['metrics.conversionsValue']);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\});
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0return Object.values(summedData);
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // output data to tab in sheet
\f1\fs24 \

\f0\fs21\fsmilli10667 function outputDataToSheet(spreadsheet, sheetName, data) \{\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0const startTime = new Date(); // Log the start time
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let sheet = spreadsheet.getSheetByName(sheetName);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0if (!sheet) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0sheet = spreadsheet.insertSheet(sheetName);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Check if data is undefined or empty, and if so, simply return\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0if (!data || data.length === 0) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0Logger.log('No data to write to ' + sheetName);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0return;\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Determine the number of columns in the data
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let numberOfColumns = Object.keys(data[0]).length;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Clear the dynamic range based on the data length
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let rangeToClear = sheet.getRange(1, 1, sheet.getMaxRows(), numberOfColumns);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0rangeToClear.clearContent();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Prepare the data for writing to the sheet
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let outputData;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0if (!Array.isArray(data[0])) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// Data is an array of objects, extract the headers and values
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const headers = Object.keys(data[0]);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const values = data.map(row => headers.map(header => row[header] !== null && row[header] !== undefined ? row[header] : ""));
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0outputData = [headers].concat(values);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\} else \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// Data is already an array of arrays
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0outputData = data;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Write all data at once to the sheet
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0const numRows = outputData.length;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0const numColumns = outputData[0].length;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0const range = sheet.getRange(1, 1, numRows, numColumns);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0range.setValues(outputData);
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0const endTime = new Date(); // Log the end time
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0//Logger.log('Duration for sheet: ' + sheetName + ' is ' + (endTime - startTime) + 'ms'); // Log the duration
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // output summary data
\f1\fs24 \

\f0\fs21\fsmilli10667 function outputSummaryToSheet(ss, sheetName, summaryData) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Get the sheet with the name `sheetName` from `ss`. Create it if it doesn't exist.
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let sheet = ss.getSheetByName(sheetName);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0if (!sheet) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0sheet = ss.insertSheet(sheetName);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0sheet.clear();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0for (let i = 0; i < summaryData.length; i++) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0sheet.getRange(i + 1, 1, 1, summaryData[i].length).setValues([summaryData[i]]);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0sheet.setFrozenRows(1);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0sheet.getRange("1:1").setFontWeight("bold");
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0sheet.getRange("1:1").setWrap(true);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0sheet.setRowHeight(1, 40);
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // create summary on one tab
\f1\fs24 \

\f0\fs21\fsmilli10667 function processSummaryData(processedCampData, processedAssetGroupData, processedDisplayData, processedDataForVideo, searchResults) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0const header = ['Date', 'Campaign Name',\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0'Camp Cost', 'Camp Conv', 'Camp Value',
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0'Shop Cost', 'Shop Conv', 'Shop Value',
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0'Disp Cost', 'Disp Conv', 'Disp Value',
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0'Video Cost', 'Video Conv', 'Video Value',
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0'Search Cost', 'Search Conv', 'Search Value',
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0'Campaign Type'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0const summaryData = \{\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Helper function to update summaryData
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0function updateSummary(row, type) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const date = row['date'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const campaignName = row['campaignName'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const campaignType = row['campaignType'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const key = `$\{date\}_$\{campaignName\}`;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (!summaryData[key]) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0summaryData[key] = \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'date': date,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'campaignName': campaignName,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'generalCost': 0, 'generalConv': 0, 'generalConvValue': 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'shoppingCost': 0, 'shoppingConv': 0, 'shoppingConvValue': 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'displayCost': 0, 'displayConv': 0, 'displayConvValue': 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'videoCost': 0, 'videoConv': 0, 'videoConvValue': 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'searchCost': 0, 'searchConv': 0, 'searchConvValue': 0,\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'campaignType': campaignType
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0summaryData[key][`$\{type\}Cost`] += row['cost'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0summaryData[key][`$\{type\}Conv`] += row['conversions'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0summaryData[key][`$\{type\}ConvValue`] += row['conversionsValue'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Process each data type
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Correctly handle shopping and Performance Max campaigns in summary data
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0processedCampData.forEach(row => \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (row['campaignType'] === 'SHOPPING') \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0updateSummary(row, 'shopping'); // Use campaign data directly for shopping campaigns
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\} else \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0updateSummary(row, 'general');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\});
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Only use asset group data for Performance Max campaigns, not for shopping campaigns
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0processedAssetGroupData.forEach(row => \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (row['campaignType'] !== 'SHOPPING') \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0updateSummary(row, 'shopping');\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// No else part needed as we're not processing shopping campaigns with asset group data
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\});
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0processedDisplayData.forEach(row => updateSummary(row, 'display'));
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0processedDataForVideo.forEach(row => updateSummary(row, 'video'));
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0searchResults.forEach(row => updateSummary(row, 'search'));
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Adjust metrics
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0Object.keys(summaryData).forEach(key => \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const data = summaryData[key];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let totalCost = data.generalCost;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let totalConv = data.generalConv;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let totalConvValue = data.generalConvValue;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// Adjust Video
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let residualCostForVideo = totalCost - data.shoppingCost;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let residualConvForVideo = totalConv - data.shoppingConv;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let residualConvValueForVideo = totalConvValue - data.shoppingConvValue;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (data.videoCost > residualCostForVideo) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0data.videoCost = residualCostForVideo;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0data.videoConv = Math.min(data.videoConv, residualConvForVideo);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0data.videoConvValue = Math.min(data.videoConvValue, residualConvValueForVideo);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// Adjust Display
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let residualCostForDisplay = totalCost - data.shoppingCost - data.videoCost;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let residualConvForDisplay = totalConv - data.shoppingConv - data.videoConv;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let residualConvValueForDisplay = totalConvValue - data.shoppingConvValue - data.videoConvValue;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (data.displayCost > residualCostForDisplay) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0data.displayCost = residualCostForDisplay;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0data.displayConv = Math.min(data.displayConv, residualConvForDisplay);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0data.displayConvValue = Math.min(data.displayConvValue, residualConvValueForDisplay);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// Adjust Search
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let residualCostForSearch = totalCost - data.shoppingCost - data.videoCost - data.displayCost;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let residualConvForSearch = totalConv - data.shoppingConv - data.videoConv - data.displayConv;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let residualConvValueForSearch = totalConvValue - data.shoppingConvValue - data.videoConvValue - data.displayConvValue;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0data.searchCost = Math.max(residualCostForSearch, 0);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0data.searchConv = Math.max(residualConvForSearch, 0);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0data.searchConvValue = Math.max(residualConvValueForSearch, 0);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\});
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Convert summaryData to an array format for output
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0const summaryArray = Object.values(summaryData).map(summaryRow => \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0return [
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0summaryRow.date,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0summaryRow.campaignName,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0summaryRow.generalCost, summaryRow.generalConv, summaryRow.generalConvValue,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0summaryRow.shoppingCost, summaryRow.shoppingConv, summaryRow.shoppingConvValue,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0summaryRow.displayCost, summaryRow.displayConv, summaryRow.displayConvValue,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0summaryRow.videoCost, summaryRow.videoConv, summaryRow.videoConvValue,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0summaryRow.searchCost, summaryRow.searchConv, summaryRow.searchConvValue,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0summaryRow.campaignType
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\});
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0return [header, ...summaryArray];
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // Process total data for 'Totals' tab
\f1\fs24 \

\f0\fs21\fsmilli10667 function processTotalData(processedCampData, processedAssetGroupData, processedDisplayData, processedDataForVideo, searchResults) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0const header = ['Campaign Name',\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0'Camp Cost', 'Camp Conv', 'Camp Value',
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0'Shop Cost', 'Shop Conv', 'Shop Value',
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0'Disp Cost', 'Disp Conv', 'Disp Value',
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0'Video Cost', 'Video Conv', 'Video Value',
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0'Search Cost', 'Search Conv', 'Search Value',
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0'Campaign Type'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0const totalData = \{\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0function updateTotal(row, type) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const campaignName = row['campaignName'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const campaignType = row['campaignType'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const key = campaignName;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (!totalData[key]) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0totalData[key] = \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'campaignName': campaignName,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'generalCost': 0, 'generalConv': 0, 'generalConvValue': 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'shoppingCost': 0, 'shoppingConv': 0, 'shoppingConvValue': 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'displayCost': 0, 'displayConv': 0, 'displayConvValue': 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'videoCost': 0, 'videoConv': 0, 'videoConvValue': 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'searchCost': 0, 'searchConv': 0, 'searchConvValue': 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'campaignType': campaignType
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0totalData[key][`$\{type\}Cost`] += row['cost'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0totalData[key][`$\{type\}Conv`] += row['conversions'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0totalData[key][`$\{type\}ConvValue`] += row['conversionsValue'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0processedCampData.forEach(row => \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0updateTotal(row, 'general');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (row['campaignType'] === 'SHOPPING') \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0updateTotal(row, 'shopping');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\});
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0processedAssetGroupData.forEach(row => \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (row['campaignType'] !== 'SHOPPING') \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0updateTotal(row, 'shopping');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\});
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0processedDataForVideo.forEach(row => \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (row['campaignType'] !== 'SHOPPING') \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0updateTotal(row, 'video');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\});
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0processedDisplayData.forEach(row => \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (row['campaignType'] !== 'SHOPPING') \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0updateTotal(row, 'display');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\});
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0searchResults.forEach(row => \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (row['campaignType'] !== 'SHOPPING') \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0updateTotal(row, 'search');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\});
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Adjust metrics for each campaign
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0Object.keys(totalData).forEach(key => \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const data = totalData[key];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let totalCost = data.generalCost;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let totalConv = data.generalConv;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let totalConvValue = data.generalConvValue;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// Adjust Video
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (data.videoCost > totalCost - data.shoppingCost) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0data.videoCost = totalCost - data.shoppingCost;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0data.videoConv = Math.min(data.videoConv, totalConv - data.shoppingConv);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0data.videoConvValue = Math.min(data.videoConvValue, totalConvValue - data.shoppingConvValue);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// Adjust Display
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let residualForDisplay = totalCost - data.shoppingCost - data.videoCost;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (data.displayCost > residualForDisplay) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0data.displayCost = residualForDisplay;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0data.displayConv = Math.min(data.displayConv, totalConv - data.shoppingConv - data.videoConv);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0data.displayConvValue = Math.min(data.displayConvValue, totalConvValue - data.shoppingConvValue - data.videoConvValue);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// Adjust Search
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let residualForSearch = totalCost - data.shoppingCost - data.videoCost - data.displayCost;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (residualForSearch < 0) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0data.searchCost = 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0data.searchConv = 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0data.searchConvValue = 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\} else \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0data.searchCost = residualForSearch;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0data.searchConv = Math.max(totalConv - data.shoppingConv - data.videoConv - data.displayConv, 0);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0data.searchConvValue = Math.max(totalConvValue - data.shoppingConvValue - data.videoConvValue - data.displayConvValue, 0);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\});
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Convert totalData to an array format
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0const totalArray = Object.values(totalData).map(totalRow => \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0return [
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0totalRow.campaignName,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0totalRow.generalCost, totalRow.generalConv, totalRow.generalConvValue,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0totalRow.shoppingCost, totalRow.shoppingConv, totalRow.shoppingConvValue,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0totalRow.displayCost, totalRow.displayConv, totalRow.displayConvValue,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0totalRow.videoCost, totalRow.videoConv, totalRow.videoConvValue,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0totalRow.searchCost, totalRow.searchConv, totalRow.searchConvValue,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0totalRow.campaignType
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\});
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0return [header, ...totalArray];
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // get product data & iterate through the three use cases
\f1\fs24 \

\f0\fs21\fsmilli10667 function fetchProductData(queryString, tCost, tRoas, outputType) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let data = [];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let aggregatedData = \{\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0const iterator = AdsApp.search(queryString);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0while (iterator.hasNext()) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const row = iterator.next();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let rowData = flattenObject(row);\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// Determine unique key based on output type
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let uniqueKey = getUniqueKey(rowData, outputType);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// Initialize aggregated data object if not present
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (!aggregatedData[uniqueKey]) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregatedData[uniqueKey] = \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'Impr': 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'Clicks': 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'Cost': 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'Conv': 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'Value': 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'Product Title': rowData['segments.productTitle'],
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'Product ID': rowData['segments.productItemId'],
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'Campaign': rowData['campaign.name'],
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'Channel': rowData['campaign.advertisingChannelType']
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// Aggregate data
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let aggData = aggregatedData[uniqueKey];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggData['Impr'] += Number(rowData['metrics.impressions']) || 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggData['Clicks'] += Number(rowData['metrics.clicks']) || 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggData['Cost'] += (Number(rowData['metrics.costMicros']) / 1e6) || 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggData['Conv'] += Number(rowData['metrics.conversions']) || 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggData['Value'] += Number(rowData['metrics.conversionsValue']) || 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Post-processing for additional fields and calculations
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0for (let key in aggregatedData) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let aggData = aggregatedData[key];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggData['ROAS'] = aggData['Cost'] > 0 ? aggData['Value'] / aggData['Cost'] : 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggData['CvR'] = aggData['Clicks'] > 0 ? aggData['Conv'] / aggData['Clicks'] : 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggData['CTR'] = aggData['Impr'] > 0 ? aggData['Clicks'] / aggData['Impr'] : 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggData['Bucket'] = determineBucket(aggData['Cost'], aggData['Conv'], aggData['ROAS'], tCost, tRoas);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// Construct base data object based on output type
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let baseDataObject = constructBaseDataObject(aggData, outputType);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0data.push(baseDataObject);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0return data;
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // helper func for product data
\f1\fs24 \

\f0\fs21\fsmilli10667 function getUniqueKey(rowData, outputType) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0switch (outputType) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0case 'prodTitle':
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0return rowData['segments.productTitle'] + '|' + rowData['campaign.name'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0case 'prodTitleID':
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0return rowData['segments.productTitle'] + '|' + rowData['segments.productItemId'] + '|' + rowData['campaign.name'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0case 'idChannel':
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0return rowData['segments.productItemId'] + '|' + rowData['campaign.advertisingChannelType'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0case 'idAccount':
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0return rowData['segments.productItemId'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0case 'prodAccount':
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0return rowData['segments.productTitle'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0default:
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0throw new Error("Invalid output type");
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // helper func for product data
\f1\fs24 \

\f0\fs21\fsmilli10667 function determineBucket(cost, conv, roas, tCost, tRoas) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0if (cost === 0) return 'zombie';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0if (conv === 0) return 'zeroconv';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0if (cost < tCost) return roas < tRoas ? 'meh' : 'flukes';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0return roas < tRoas ? 'costly' : 'profitable';
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // helper func for product data
\f1\fs24 \

\f0\fs21\fsmilli10667 function constructBaseDataObject(aggData, outputType) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let baseDataObject = \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'Product Title': aggData['Product Title'],
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'Product ID': aggData['Product ID'],
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'Impr': aggData['Impr'],
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'Clicks': aggData['Clicks'],
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'Cost': aggData['Cost'],
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'Conv': aggData['Conv'],
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'Value': aggData['Value'],
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'CTR': aggData['CTR'],
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'ROAS': aggData['ROAS'],
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'CvR': aggData['CvR'],
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'Bucket': aggData['Bucket'],
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'Campaign': aggData['Campaign'],
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'Channel': aggData['Channel']
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\};
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0switch (outputType) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0case 'prodTitle': // Remove the Product ID for 'prodTitle' output type
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0delete baseDataObject['Product ID'];\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0break;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0case 'prodAccount': // delete ID, campaign & channel - to aggregate across whole account
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0delete baseDataObject['Product ID'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0delete baseDataObject['Campaign'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0delete baseDataObject['Channel'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0break;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0case 'idChannel':
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0baseDataObject = \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'ID': aggData['Product ID'],
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'Bucket': aggData['Bucket'],
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'Channel': aggData['Channel']
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0break;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0case 'idAccount':
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0baseDataObject = \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'ID': aggData['Product ID'],
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'Bucket': aggData['Bucket']
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0break;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0return baseDataObject;
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // helper func for logging time & settings - no data
\f1\fs24 \

\f0\fs21\fsmilli10667 function log(ss, duration, s, ident) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let newRow = [new Date(), duration, s.numberOfDays, s.tCost, s.tRoas, s.brandTerm, ident,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0s.fromDate, s.toDate, s.lotsProducts, s.turnonTitle, s.turnonID, s.turnonTitleAccount,\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0s.campFilter, \'a0 \'a0 s.productDays, s.turnonPlace, s.turnonChange, s.turnonOpenai, s.model, s.version];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0logUrl = ss.getRangeByName('u').getValue();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0[SpreadsheetApp.openByUrl(logUrl), ss].map(s => s.getSheetByName('log')).forEach(sheet => sheet.appendRow(newRow));
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // update variables from the settings tab
\f1\fs24 \

\f0\fs21\fsmilli10667 function updateVariablesFromSheet(ss, defaultSettings) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0const settingsSheetName = 'Settings'; // The name of the settings sheet in the client's spreadsheet
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0const advancedSettingsSheetName = 'Advanced'; // The name of the advanced settings sheet in the client's spreadsheet
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// add advanced setting defaults\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0defaultSettings.fromDate = undefined
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0defaultSettings.toDate = undefined
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0defaultSettings.lotsProducts = 0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0defaultSettings.turnonTitle = false
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0defaultSettings.turnonID = false
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0defaultSettings.turnonTitleAccount = false
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0defaultSettings.campFilter = ''
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0defaultSettings.turnonPlace = false
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0defaultSettings.turnonChange = false
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0defaultSettings.turnonOpenai = false
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0defaultSettings.apiKey = ''
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0defaultSettings.model = ''
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0defaultSettings.version = 'v56'
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0try \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const settingsSheet = ss.getSheetByName(settingsSheetName);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const advancedSheet = ss.getSheetByName(advancedSettingsSheetName);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (!settingsSheet || !advancedSheet) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0console.error('One or both settings tabs not found in client spreadsheet.');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0return defaultSettings;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// Initialize an object to hold the updated settings
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let updatedSettings = \{...defaultSettings\};
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// Helper function to safely get range values
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0function getRangeValue(sheet, rangeName) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0try \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const range = sheet.getRange(rangeName);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const value = range.getDisplayValue();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0return value !== '' ? value : null;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\} catch (e) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0console.error(`Range $\{rangeName\} not found in sheet $\{sheet.getName()\}:`, e);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0return null;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// Settings that are numbers or strings
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const settingsRangeNames = ['numberOfDays', 'tCost', 'tRoas', 'brandTerm'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0settingsRangeNames.forEach(settingKey => \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const value = getRangeValue(settingsSheet, settingKey);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0updatedSettings[settingKey] = value !== null ? (isNaN(value) ? value.trim() : Number(value)) : defaultSettings[settingKey];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\});
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// Advanced settings that are dates, booleans, or numbers
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const advancedSettingsRangeNames = ['fromDate', 'toDate', 'lotsProducts', 'turnonTitle', 'turnonID', 'turnonTitleAccount',\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0'campFilter', 'productDays', 'turnonPlace', 'turnonChange', 'turnonOpenai', 'apiKey', 'model'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0advancedSettingsRangeNames.forEach(settingKey => \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const value = getRangeValue(advancedSheet, settingKey);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (value !== null) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (settingKey === 'fromDate' || settingKey === 'toDate') \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0updatedSettings[settingKey] = /^\\d\{2\}\\/\\d\{2\}\\/\\d\{4\}$/.test(value) ? value : defaultSettings[settingKey];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\} else if (settingKey === 'turnonTitle' || settingKey === 'turnonID' || settingKey === 'turnonTitleAccount' ||\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0settingKey === 'turnonPlace' || settingKey === 'turnonChange' || settingKey === 'turnonOpenai') \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0updatedSettings[settingKey] = value.toLowerCase() === 'true';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\} else if (settingKey === 'lotsProducts') \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0updatedSettings[settingKey] = isNaN(value) ? defaultSettings[settingKey] : Number(value);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\} else \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0updatedSettings[settingKey] = value.trim();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\} else \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0updatedSettings[settingKey] = defaultSettings[settingKey];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\});
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0return updatedSettings;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\} catch (e) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0console.error("Error in updateVariablesFromSheet:", e);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0return defaultSettings;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // extract search terms
\f1\fs24 \

\f0\fs21\fsmilli10667 function extractTerms(ss, mainDateRange, brands, tol) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Extract Campaign IDs with Status not 'REMOVED' and impressions > 0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let campaignIdsQuery = AdsApp.report(`
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0SELECT campaign.id, campaign.name, metrics.clicks,\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0metrics.impressions, metrics.conversions, metrics.conversions_value
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0FROM campaign\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0WHERE campaign.status != 'REMOVED'\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0AND campaign.advertising_channel_type = "PERFORMANCE_MAX"\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0AND metrics.impressions > 0 AND $\{mainDateRange\}\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0ORDER BY metrics.conversions DESC `
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let rows = campaignIdsQuery.rows();\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let allSearchTerms = [['Campaign Name', 'Campaign ID', 'Category Label', 'Clicks', 'Impr', 'Conv', 'Value', 'Bucket', 'Distance']];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0while (rows.hasNext()) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let row = rows.next();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// Fetch search categories and related metrics for the campaign ordered by conversions descending
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let query = AdsApp.report(`\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0SELECT campaign_search_term_insight.category_label, metrics.clicks,\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0metrics.impressions, metrics.conversions, metrics.conversions_value\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0FROM campaign_search_term_insight\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0WHERE $\{mainDateRange\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0AND campaign_search_term_insight.campaign_id = $\{row['campaign.id']\}\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0ORDER BY metrics.impressions DESC `
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let searchTermRows = query.rows();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0while (searchTermRows.hasNext()) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let searchTermRow = searchTermRows.next();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let term = searchTermRow['campaign_search_term_insight.category_label'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let \{ bucket, distance \} = determineBucketAndDistance(term, brands, tol);\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0allSearchTerms.push([row['campaign.name'], row['campaign.id'],\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0term,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0searchTermRow['metrics.clicks'],\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0searchTermRow['metrics.impressions'],\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0searchTermRow['metrics.conversions'],\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0searchTermRow['metrics.conversions_value'],
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0bucket,\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0distance]);\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Write all search terms to the 'terms' sheet\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let termsSheet = ss.getSheetByName('terms') ? ss.getSheetByName('terms').clear() : ss.insertSheet('terms');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0termsSheet.getRange(1, 1, allSearchTerms.length, allSearchTerms[0].length).setValues(allSearchTerms);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Aggregate terms and write to the 'totalTerms' sheet
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0aggregateTerms(allSearchTerms, ss);
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // aggregate search terms
\f1\fs24 \

\f0\fs21\fsmilli10667 function aggregateTerms(searchTerms, ss) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let aggregated = \{\}; // \{ term: \{ clicks: 0, impressions: 0, conversions: 0, conversionValue: 0, bucket: '', distance: 0 \}, ... \}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0for (let i = 1; i < searchTerms.length; i++) \{ // Start from 1 to skip headers
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let term = searchTerms[i][2] || 'blank'; // Use 'blank' for empty search terms
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (!aggregated[term]) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregated[term] = \{\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0clicks: 0,\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0impressions: 0,\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0conversions: 0,\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0conversionValue: 0,\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0bucket: searchTerms[i][7],\'a0 // Assuming bucket is in the 8th position of your array
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0distance: searchTerms[i][8]\'a0 // Assuming distance is in the 9th position of your array
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregated[term].clicks += Number(searchTerms[i][3]);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregated[term].impressions += Number(searchTerms[i][4]);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregated[term].conversions += Number(searchTerms[i][5]);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregated[term].conversionValue += Number(searchTerms[i][6]);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// Assuming that the bucket and distance are the same for all instances of a term,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// we don't aggregate them but just take the value from the first instance.
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let aggregatedArray = [['Category Label', 'Clicks', 'Impressions', 'Conversions', 'Conversion Value', 'Bucket', 'Distance']]; // Header row
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0for (let term in aggregated) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregatedArray.push([
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0term,\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregated[term].clicks,\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregated[term].impressions,\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregated[term].conversions,\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregated[term].conversionValue,\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregated[term].bucket,\'a0 // Adding bucket to output
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregated[term].distance\'a0 // Adding distance to output
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0]);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let header = aggregatedArray.shift(); // Remove the header before sorting
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Sort by impressions descending
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0aggregatedArray.sort((a, b) => b[2] - a[2]);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0aggregatedArray.unshift(header); // Prepend the header back to the top
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Write aggregated data to the 'totalTerms' sheet
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let totalTermsSheet = ss.getSheetByName('totalTerms') ? ss.getSheetByName('totalTerms').clear() : ss.insertSheet('totalTerms');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0totalTermsSheet.getRange(1, 1, aggregatedArray.length, aggregatedArray[0].length).setValues(aggregatedArray);
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // Helper function to determine the bucket for a term and calculate Levenshtein distance
\f1\fs24 \

\f0\fs21\fsmilli10667 function determineBucketAndDistance(term, brandTerms, tolerance) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0const brandTermsArray = (typeof brandTerms === 'string') ? brandTerms.split(',').map(term => term.trim()).filter(term => term) : [];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0if (brandTermsArray.length === 0) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0brandTermsArray.push('no brand has been entered on the sheet');\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0const lowerCaseTerm = term.toLowerCase();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0if (lowerCaseTerm === 'blank' || lowerCaseTerm === '') \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0return \{ bucket: 'blank', distance: '' \};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// First, check for an exact match
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0for (const brandTerm of brandTermsArray) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const lowerCaseBrandTerm = brandTerm.toLowerCase();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (lowerCaseTerm === lowerCaseBrandTerm) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0return \{ bucket: 'brand', distance: 0 \};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// If no exact match is found, check for close matches
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0for (const brandTerm of brandTermsArray) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const lowerCaseBrandTerm = brandTerm.toLowerCase();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0const distance = levenshtein(lowerCaseTerm, lowerCaseBrandTerm);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (lowerCaseTerm.includes(lowerCaseBrandTerm) || distance <= tolerance) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0return \{ bucket: 'close-brand', distance \};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// If no close match is found, it's a non-brand
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0return \{ bucket: 'non-brand', distance: '' \};
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // Function to calculate Levenshtein distance between two strings
\f1\fs24 \

\f0\fs21\fsmilli10667 function levenshtein(a, b) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let tmp;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0if (a.length === 0) \{ return b.length; \}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0if (b.length === 0) \{ return a.length; \}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0if (a.length > b.length) \{ tmp = a; a = b; b = tmp; \}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let i, j, res, alen = a.length, blen = b.length, row = Array(alen);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0for (i = 0; i <= alen; i++) \{ row[i] = i; \}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0for (i = 1; i <= blen; i++) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0res = i;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0for (j = 1; j <= alen; j++) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0tmp = row[j - 1];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0row[j - 1] = res;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0res = b[i - 1] === a[j - 1] ? tmp : Math.min(tmp + 1, Math.min(res + 1, row[j] + 1));
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0return res;
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // find single word nGrams
\f1\fs24 \

\f0\fs21\fsmilli10667 function extractNGrams(ss, mainDateRange) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Extract Campaign IDs with Status not 'REMOVED' and impressions > 0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let campaignIdsQuery = AdsApp.report(`SELECT campaign.id, campaign.name, metrics.clicks, metrics.impressions, metrics.conversions, metrics.conversions_value
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0FROM campaign WHERE campaign.status != 'REMOVED' AND metrics.impressions > 0 AND $\{mainDateRange\} ORDER BY metrics.conversions DESC `
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let rows = campaignIdsQuery.rows();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let campaignNGrams = \{\}; // To store nGram data per campaign
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0while (rows.hasNext()) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let row = rows.next();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let campaignId = row['campaign.id'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let campaignName = row['campaign.name'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// Fetch search terms and related metrics for the campaign ordered by conversions descending
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let query = AdsApp.report(` SELECT campaign_search_term_insight.category_label, metrics.clicks, metrics.impressions,\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0metrics.conversions, metrics.conversions_value\'a0 FROM campaign_search_term_insight WHERE $\{mainDateRange\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0AND campaign_search_term_insight.campaign_id = $\{campaignId\} AND metrics.impressions > 0 ORDER BY metrics.impressions DESC `
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let searchTermRows = query.rows();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0while (searchTermRows.hasNext()) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let searchTermRow = searchTermRows.next();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let terms = searchTermRow['campaign_search_term_insight.category_label'].split(' '); // Split the category label into individual words
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0terms.forEach((term) => \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let key = campaignId + '_' + term; // Unique key for each nGram within a campaign
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (!campaignNGrams[key]) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0campaignNGrams[key] = \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0campaignName: campaignName,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0campaignId: campaignId,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0nGram: term,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0clicks: 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0impressions: 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0conversions: 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0conversionValue: 0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// Aggregate the metrics for the nGram within the campaign
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0campaignNGrams[key].clicks += Number(searchTermRow['metrics.clicks']);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0campaignNGrams[key].impressions += Number(searchTermRow['metrics.impressions']);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0campaignNGrams[key].conversions += Number(searchTermRow['metrics.conversions']);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0campaignNGrams[key].conversionValue += Number(searchTermRow['metrics.conversions_value']);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\});
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Convert the aggregated object to an array format suitable for the spreadsheet
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let allNGrams = [['Campaign Name', 'Campaign ID', 'nGram', 'Clicks', 'Impressions', 'Conversions', 'Conversion Value']];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0for (let key in campaignNGrams) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let item = campaignNGrams[key];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0allNGrams.push([item.campaignName, item.campaignId, item.nGram, item.clicks, item.impressions, item.conversions, item.conversionValue]);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Write all nGrams to the 't2' sheet
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let nGramsSheet = ss.getSheetByName('t2') ? ss.getSheetByName('t2').clear() : ss.insertSheet('t2');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0nGramsSheet.getRange(1, 1, allNGrams.length, allNGrams[0].length).setValues(allNGrams);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Continue with aggregation for 'total2' as before
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0aggregateNGrams(allNGrams, ss);
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // aggregate the nGrams
\f1\fs24 \

\f0\fs21\fsmilli10667 function aggregateNGrams(nGrams, ss) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Similar structure to aggregateTerms, but we're aggregating nGrams
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let aggregated = \{\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0for (let i = 1; i < nGrams.length; i++) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let nGram = nGrams[i][2] || 'blank';
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (!aggregated[nGram]) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregated[nGram] = \{ clicks: 0, impressions: 0, conversions: 0, conversionValue: 0 \};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregated[nGram].clicks += Number(nGrams[i][3]);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregated[nGram].impressions += Number(nGrams[i][4]);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregated[nGram].conversions += Number(nGrams[i][5]);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregated[nGram].conversionValue += Number(nGrams[i][6]);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let aggregatedArray = [['nGram', 'Clicks', 'Impressions', 'Conversions', 'Conversion Value']];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0for (let nGram in aggregated) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0aggregatedArray.push([nGram, aggregated[nGram].clicks, aggregated[nGram].impressions, aggregated[nGram].conversions, aggregated[nGram].conversionValue]);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let header = aggregatedArray.shift();
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0aggregatedArray.sort((a, b) => b[2] - a[2]); // Sort by impressions
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0aggregatedArray.unshift(header);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Write aggregated data to the 'total2' sheet
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let totalNGramsSheet = ss.getSheetByName('total2') ? ss.getSheetByName('total2').clear() : ss.insertSheet('total2');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0totalNGramsSheet.getRange(1, 1, aggregatedArray.length, aggregatedArray[0].length).setValues(aggregatedArray);
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // output placement data
\f1\fs24 \

\f0\fs21\fsmilli10667 function outputPlaceToSheet(ss, data, numDays) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let sheet = ss.getSheetByName('place');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0if (!sheet) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0console.log('Sheet not found.');
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0return;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // Clear the existing data in all but the last column
\f1\fs24 \

\f0\fs21\fsmilli10667 let lastRow\'a0 \'a0 = sheet.getLastRow();
\f1\fs24 \

\f0\fs21\fsmilli10667 let lastColumn = sheet.getLastColumn();
\f1\fs24 \

\f0\fs21\fsmilli10667 if (lastRow > 1) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Adjusted to clear sheet & leave headers (& leave last 2 columns)
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let range = sheet.getRange(2, 1, lastRow - 1, lastColumn - 2);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0range.clearContent();
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 let costLimit = numDays / 10\'a0 // assume must spend at least 10c/day
\f1\fs24 \

\f0\fs21\fsmilli10667 let imprLimit = numDays * 100 // assume must get at least 100 impr/day - maybe use just for titles
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Convert the aggregated object into an array of rows and filter in one step
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0let rows = [];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0for (let placement in data) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let item = data[placement];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (item.cost >= costLimit && placement !== 'undefined' && placement) \{ // && item.impr >= imprLimit
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0rows.push([
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0item.placementType, placement, item.impr, item.interact,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0item.views, item.cost, item.conv, item.value,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0item.allConv, item.allValue,\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0item.cpv, item.intRate, item.viewRate,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0item.cvr, item.roas, item.cpa, item.cpi,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0item.allCvr, item.allCpa, item.allCpi
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0]);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0// Write the filtered data to the sheet starting from the second row
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0if (rows.length > 0) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // aggregate placement data
\f1\fs24 \

\f0\fs21\fsmilli10667 function aggMet(rows) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0return Array.from(rows).reduce((acc, row) => \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let p \'a0 \'a0 \'a0 \'a0 \'a0 \'a0 = row['detail_placement_view.group_placement_target_url'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let placementType = row['detail_placement_view.placement_type'];
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let impr\'a0 \'a0 \'a0 \'a0 \'a0 = parseInt(row['metrics.impressions'], 10) \'a0 \'a0 \'a0 \'a0 || 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let interact\'a0 \'a0 \'a0 = parseInt(row['metrics.interactions'], 10)\'a0 \'a0 \'a0 \'a0 || 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let views \'a0 \'a0 \'a0 \'a0 = parseInt(row['metrics.video_views'], 10) \'a0 \'a0 \'a0 \'a0 || 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let cost\'a0 \'a0 \'a0 \'a0 \'a0 = parseInt(row['metrics.cost_micros'], 10) / 1e6 \'a0 || 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let conv\'a0 \'a0 \'a0 \'a0 \'a0 = parseFloat(row['metrics.conversions']) \'a0 \'a0 \'a0 \'a0 \'a0 || 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let value \'a0 \'a0 \'a0 \'a0 = parseFloat(row['metrics.conversions_value']) \'a0 \'a0 || 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let allConv \'a0 \'a0 \'a0 = parseFloat(row['metrics.all_conversions']) \'a0 \'a0 \'a0 || 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0let allValue\'a0 \'a0 \'a0 = parseFloat(row['metrics.all_conversions_value']) || 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0if (!acc[p]) \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0acc[p] = \{\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0placementType, impr, interact, views, cost, conv, value, allConv, allValue,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0cpv:\'a0 \'a0 \'a0 views > 0\'a0 \'a0 ? cost / views : 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0intRate:\'a0 impr > 0 \'a0 \'a0 ? interact / impr : 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0viewRate: impr > 0 \'a0 \'a0 ? views / impr : 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0cvr:\'a0 \'a0 \'a0 interact > 0 ? conv / interact : 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0roas: \'a0 \'a0 cost > 0 \'a0 \'a0 ? value / cost : 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0cpa:\'a0 \'a0 \'a0 conv > 0 \'a0 \'a0 ? cost / conv : 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0cpi:\'a0 \'a0 \'a0 impr > 0 \'a0 \'a0 ? (conv / impr) * 1000 : 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0allCvr: \'a0 interact > 0 ? allConv / interact : 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0allCpa: \'a0 allConv > 0\'a0 ? cost / allConv : 0,
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0allCpi: \'a0 impr > 0 \'a0 \'a0 ? (allConv / impr) * 1000 : 0
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\};
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\} else \{
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0acc[p].impr \'a0 \'a0 += impr;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0acc[p].interact += interact;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0acc[p].views\'a0 \'a0 += views;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0acc[p].cost \'a0 \'a0 += cost;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0acc[p].conv \'a0 \'a0 += conv;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0acc[p].value\'a0 \'a0 += value;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0acc[p].allConv\'a0 += allConv;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0acc[p].allValue += allValue;
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0// calculate additional metrics
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0acc[p].cpv \'a0 \'a0 \'a0 = acc[p].views\'a0 \'a0 > 0 ? acc[p].cost / acc[p].views : 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0acc[p].intRate \'a0 = acc[p].impr \'a0 \'a0 > 0 ? acc[p].interact / acc[p].impr : 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0acc[p].viewRate\'a0 = acc[p].impr \'a0 \'a0 > 0 ? acc[p].views / acc[p].impr : 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0acc[p].cvr \'a0 \'a0 \'a0 = acc[p].interact > 0 ? acc[p].conv / acc[p].interact : 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0acc[p].roas\'a0 \'a0 \'a0 = acc[p].cost \'a0 \'a0 > 0 ? acc[p].value / acc[p].cost : 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0acc[p].cpa \'a0 \'a0 \'a0 = acc[p].conv \'a0 \'a0 > 0 ? acc[p].cost / acc[p].conv : 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0acc[p].cpi \'a0 \'a0 \'a0 = acc[p].impr \'a0 \'a0 > 0 ? (acc[p].conv / acc[p].impr) * 1000 : 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0acc[p].allCvr\'a0 \'a0 = acc[p].interact > 0 ? acc[p].allConv / acc[p].interact : 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0acc[p].allCpa\'a0 \'a0 = acc[p].allConv\'a0 > 0 ? acc[p].cost / acc[p].allConv : 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0acc[p].allCpi\'a0 \'a0 = acc[p].impr \'a0 \'a0 > 0 ? (acc[p].allConv / acc[p].impr) * 1000 : 0;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0\}
\f1\fs24 \
\

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\'a0\'a0\'a0\'a0return acc;
\f1\fs24 \

\f0\fs21\fsmilli10667 \'a0\'a0\'a0\'a0\}, \{\});
\f1\fs24 \

\f0\fs21\fsmilli10667 \}
\f1\fs24 \
\
\

\f0\fs21\fsmilli10667 /*
\f1\fs24 \

\f0\fs21\fsmilli10667 DISCLAIMER\'a0 -\'a0 PLEASE READ CAREFULLY BEFORE USING THIS SCRIPT
\f1\fs24 \
\

\f0\fs21\fsmilli10667 Fair Use: This script is provided for the sole use of the entity (business, agency, or individual) to which it is licensed.\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 While you are encouraged to use and benefit from this script, you must do so within the confines of this agreement.
\f1\fs24 \
\

\f0\fs21\fsmilli10667 Copyright: All rights, including copyright, in this script are owned by or licensed to Off Rhodes Pty Ltd t/a Mike Rhodes Ideas.\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 Reproducing, distributing, or selling any version of this script, whether modified or unmodified, without proper authorization is strictly prohibited.
\f1\fs24 \
\

\f0\fs21\fsmilli10667 License Requirement: A separate license must be purchased for each legal entity that wishes to use this script.\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 For example, if you own multiple businesses or agencies, each business or agency can use this script under one license.\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 However, if you are part of a holding group or conglomerate with multiple separate entities, each entity must purchase its own license for use.
\f1\fs24 \
\

\f0\fs21\fsmilli10667 Code of Honour: This script is offered under a code of honour.\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 We trust in the integrity of our users to adhere to the terms of this agreement and to do the right thing.\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 Your honour and professionalism in respecting these terms not only supports the creator but also fosters a community of trust and respect.
\f1\fs24 \
\

\f0\fs21\fsmilli10667 Limitations & Liabilities: Off Rhodes Pty Ltd t/a Mike Rhodes Ideas does not guarantee that this script will be error-free\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 or will meet your specific requirements. We are not responsible for any damages or losses that might arise from using this script.\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 Always back up your data and test the script in a safe environment before deploying it in a production setting.
\f1\fs24 \
\

\f0\fs21\fsmilli10667 The script does not make any changes to your account or data.\'a0
\f1\fs24 \

\f0\fs21\fsmilli10667 It only reads data from your account and writes it to your spreadsheet.
\f1\fs24 \

\f0\fs21\fsmilli10667 However if you choose to use the data on the ID tab in a supplemental data feed in your GMC account, you do so at your own risk.
\f1\fs24 \
\

\f0\fs21\fsmilli10667 By using this script, you acknowledge that you have read, understood, and agree to be bound by the terms of this license agreement.
\f1\fs24 \

\f0\fs21\fsmilli10667 If you do not agree with these terms, do not use this script.
\f1\fs24 \

\f0\fs21\fsmilli10667 */
\f1\fs24 \
\
\
\
\
\

\f0\fs21\fsmilli10667 // PS you're awesome! Thanks for using this script.
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // Now hit preview (or run) and let's get started!
\f1\fs24 \
\

\f0\fs21\fsmilli10667 // Click the Logs -----\uc0\u8628 
\f1\fs24 \
\
}