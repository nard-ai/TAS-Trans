Attribute VB_Name = "PublicVariable"
Public StrDbpath As String
Public StrQuery As String

Public LoginName As String  ' login username
Public LoginPass As String  ' login password
Public LoginTime As String  ' Login Time
Public LogoutTime As String ' Logout Time
Public AccountType As String ' account type

Public AddRights As Boolean
Public EditRights As Boolean
Public DeleteRights As Boolean

Public LoadPictureFlag As Boolean

Public SLWithPay As String      ' sick leave M,P,S
Public VLWithPay As String      ' vacation L
Public HalfWithPay As String    ' Half day Y .5

Public StrDbHRpath As String
Public FMSFlag As Boolean   ' tag whether for FMS or not

Public sqlString As String

Public CompanyName As String
Public CompanyCode As String
Public CompanyShort As String
Public CurrentClientCode As Integer ' use to save the current client code for filtering

Public cnnstr As String
Public CnnstrHR As String

Public SuperUser As Boolean
Public AdminRights As Boolean
Public operatorID As Integer
Public cnn As ADODB.Connection
Public cnnHR As ADODB.Connection
Public ReportConnection As String

Public CnnstrCopyDestination As String
Public destinationCompanyCode As String

Public cancelRecontract As Boolean
Public TempNewID As String
Public reContract As Boolean

Public ClientSeriesNo As Integer    ' series number per client

Public QryHrHiredREport As String
Public logoPath As String

Public CutOffPeriod As String

Public SelectCompanyFlag As Boolean ' use for PhilHealth Report
Public PHealthFlag As Boolean ' use whether selection is SSS or PhilHeatlh
Public ReportedFlag As Integer
Public R1AFlag As Boolean
Public TimeKeepFlag As Boolean

Public SSSGroupCode As String
Public PHGroupCode As String
Public HDMFGroupCode As String

Public gGrossBill As Double ' used by billing total amount

Public RestDayWorkedOTRate As Double   ' 30%
Public SpecialWorkedDayOffRate As Double   ' 50%
Public LegalHolidayWithPremiumRate As Double   ' 100%
Public LegalHolidayWorkedDayoffRate As Double  ' 160%
Public WorkHrsPerDay As Double
Public OverTimeRate As Double
Public ExcessSplOTRate As Double
Public ExcessDayoffSplOTRate As Double
Public ExcessLglHolOTRate As Double
Public ExcessSunLglOTRate As Double
Public NightDiffRate As Double
Public AgencyRate As Double
Public VatPercentRate As Double
Public BillingTemplate As String

Public WithHazardPay As Boolean

