Option Explicit

Dim MyMsgBox
Set MyMsgBox = DotNetFactory.CreateInstance("System.Windows.Forms.MessageBox", "System.Windows.Forms")
'MyMsgBox.Show  "Start", "START"

Dim apiUser, apiSecret, octaneUrl
Dim sharedSpaceId, workspaceId, runId
 
apiUser = Parameter("aApiUser")
apiSecret = Parameter("aApiSecret")
octaneUrl = Parameter("aOctaneUrl")
sharedSpaceId = CInt(Parameter("aOctaneSpaceId"))
workspaceId = CInt(Parameter("aOctaneWorkspaceId"))
runId = Parameter("aRunId")

Dim OctaneUtil
Set OctaneUtil = DotNetFactory.CreateInstance("OctaneSdkWrapper.OctaneUtil", "OctaneSdkWrapper", apiUser, apiSecret, octaneUrl, sharedSpaceId, workspaceId, runId)

Dim i, element, attachmentsName, attachmentsList

'Get Run Data
Dim run
Set run = OctaneUtil.GetRun()
MyMsgBox.Show "ID: " + OctaneUtil.GetRunId() + vbCrLf + "SubType1: " + run.SubType + vbCrLf +"SubType2: " + run.GetValue("subtype"), "Run Data"

'Get Test Data
Set attachmentsList = OctaneUtil.GetTestAttachments()
attachmentsName = ""
For i = 0 To attachmentsList.BaseEntities.Count - 1
	Set element = attachmentsList.BaseEntities.Item(CInt(i))
	If (Len(attachmentsName) > 0) Then
		attachmentsName = attachmentsName + ", "
	End If
	attachmentsName = attachmentsName + element.Name
Next
MyMsgBox.Show "ID: " + OctaneUtil.GetTestId() + vbCrLf +"Name: " + OctaneUtil.GetTestName() + vbCrLf + "Attachments: " + attachmentsName, "Test Data"

'Get TestSuite Data
Set attachmentsList = OctaneUtil.GetTestSuiteAttachments()
attachmentsName = ""
For i = 0 To attachmentsList.BaseEntities.Count - 1
	Set element = attachmentsList.BaseEntities.Item(CInt(i))
	If (Len(attachmentsName) > 0) Then
		attachmentsName = attachmentsName + ", "
	End If
	attachmentsName = attachmentsName + element.Name
Next
MyMsgBox.Show "ID: " + OctaneUtil.GetTestSuiteId() + vbCrLf +"Name: " + OctaneUtil.GetTestSuiteName() + vbCrLf + "Attachments: " + attachmentsName, "TestSuite Data"

'Get RunSuite Data
MyMsgBox.Show "ID: " + OctaneUtil.GetRunSuiteId() + vbCrLf +"Name: " + OctaneUtil.GetRunSuiteName(), "RunSuite Data"

'Get RunBy Data
MyMsgBox.Show "ID: " + OctaneUtil.GetRunById() + vbCrLf +"Name: " + OctaneUtil.GetRunByName(), "RunBy Data"



