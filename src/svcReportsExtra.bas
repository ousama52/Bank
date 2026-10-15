Attribute VB_Name = "svcReportsExtra"
Option Explicit

' Reports of Stages 2 to 8 (called by svcReports.RunAll). Each part is skipped when its tables do not exist.
Public Sub RunExtraReports(ByVal d As Date, ByVal inp As Object)
    svcReportsCosts.RunCostReports d, inp
    svcReportsLoans.BuildAll d, inp
    svcReportsCollateral.RunCollateralReports d, inp
    svcReportsStaff.RunStaffReports d, inp
    svcReportsFx.RunFxReports d, inp
    svcReportsInvest.RunInvestReports d, inp
    svcReportsGov.RunGovReports d, inp
End Sub
