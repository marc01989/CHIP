VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Sys Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


'Created by:             Marco Caruso: Quality Audit Dept, Technical Team - carusomj@upmc.edu
'
'
'Modified Date:          10/21/2020
'Version Go Live:        7/13/2020
'Purpose:                To log, quantify and audit data for CHIP enrollment process
'
'
'
'                       Current Version = 2.1
'
'
'                       Change Log
'                       1.0:
'                           Stable version
'                           This is the go live verion.
'
'                       1.1
'                           1/28/2020
'                           Added warning that's triggered with leaving a record that has unsaved changes. The SetListeners() function is called on form open to add a
'                           MarkChanged() function to each control's AfterUpdate event. After a change is made to contents of textbox/combobox, MarkChanged() sets global
'                           variable bolChanged = true. When user navagates away while bolChanged = true, they will be warned to save record
'
'                           Also decresed margins between form control make everything fit within screen per Annette's recommendation in 1/23/2020 email
'
'                       1.2
'                           4/14/2020
'                           Updates to Error Assessment reports:
'                           Per Beth wolf's reequest 7495:
'                           Error assessment reports now only pull errors, they prev pulled errors and where [comments] IS NOT NULL
'                           I added field [If no updates performed by Enrollment Rep, was comment added to Caps?] to both reports. This field was already on detail reports.
'                           I removed fields [MCI Clearance Correct] and [Do we have the signature page]. These are no longer on entry forms.
'                           Record source queries for Error Assessment reports are qryErrorAssessApp & qryErrorAssessRen
'                           Had to make new record source queries for Detail reports bc they used to share with Error Assessment reports - qryDetailReportApp & qryDetailReportRen
'
'                       1.3
'                           5/12/2020
'                           Added [week_of] to error assessment reports
'                           Added Custom Search button to Apps and Rens forms. Button opens a pop up window that allows custom search values instead of just searching week_of
'                           Custom search form works by passing the name of sender form ('Apps' or 'Rens') to module then constructing sql string from user input, then sending results
'                               back to main form
'                           Also rearranged visual layout of controls at top of Apps and Rens forms to make more room for Custom Search
'                           Updated the tables that feed into Error Assessmnet reports so that they no longer overlap with detail reports and throw errors
'
'                       2.0
'                           7/13/2020
'                           Added COB reviews to database: includes entry form [COB Review], and reports [COB Audit Findings Detail Report] and [COB Audit Specialist Accuracy Summary]
'                           Also added [Admin Reporting] section to reports for Root Cause exports and high level CHIP enrollment rep report as requested by Annette on 6/22/20 demo meeting
'                           Added version # to error catching
'                           Added embedded inventory reports to report form
'
'                       2.1
'                           10/21/2020
'                           In Report Screen - added [COB Root Cause Export] and [CHIP Root Cause Export] reports
'                           These reports are exported via Excel to cloud based folders. Report layouts match templated provided by Evens and Gov Products team
'                           Reports take date parameters and capture all errors for given timeframe
'
'                       2.2
'                           12/09/2020
'                           In Report Screen - for [COB Root Cause Export] and [CHIP Root Cause Export] reports
'                           Updated the fields on the export reports to match those provided by Beth's team


Private Sub Label0_Click()
        CreateObject("Shell.application").Open "https://www.youtube.com/watch?v=60J6OGgPFAw"
End Sub
