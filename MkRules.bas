Sub ExecRulesSubjectFolder()
'CHANGE CODES FROM HERE
'///////////////////////////////////////////////////////////////////////////////////////////

'Documentation officielle de Microsoft: https://docs.microsoft.com/en-us/office/vba/outlook/How-to/Rules/create-a-rule-to-move-specific-e-mails-to-a-folder

'USE IF ONE SUBJECT
Call SetRulesSubjectFolder("RULE1", "APPLIS", "TANDEM", "TANDEM", "info@macro.fr", "info2@macro.fr")

'DO NOT CHANGE CODES
AFTER '///////////////////////////////////////////////////////////////////////////////////////////
End Sub



Sub SetRulesSubjectFolder2Subjects(strRule As String, strFolder As String, strFolder2 As String, strSubj As String, strSubj2 As String, strEmail1 As String, strEmail2 As String)
Dim colRules As Outlook.Rules
Dim oRule As Outlook.Rule
Dim colRuleActions As Outlook.RuleActions
Dim oMoveRuleAction As Outlook.MoveOrCopyRuleAction
Dim oRuleCondition As Outlook.RuleConditions
Dim oInbox As Outlook.Folder
Dim oMoveTarget As Outlook.Folder
'Specify target folder for rule move action & 'Assume that target folder already exists
Set oMoveTarget = Application.Session.GetDefaultFolder(olFolderInbox).Folders.Item(strFolder).Folders.Item(strFolder2)
'Get Rules from Session.DefaultStore object
Set colRules = Application.Session.DefaultStore.GetRules()
'Create the rule by adding a Receive Rule to Rules collection
Set oRule = colRules.Create(strRule, olRuleReceive)
'Set parameters of the rule
Set oRuleCondition = oRule.Conditions

'FIRST CONDITION
With oRuleCondition.Subject
.Enabled = True
.Text = Array(strSubj, strSubj2)
End With

'SECOND CONDITION
With oRuleCondition.SenderAddress
.Enabled = True
.Address = Array(strEmail1, strEmail2) '<--------HERE info@macro.fr; info2@macro.fr
End With

'Specify the action in a MoveOrCopyRuleAction object & 'Action is to move the message to the target folder
Set oMoveRuleAction = oRule.Actions.MoveToFolder
With oMoveRuleAction
.Enabled = True
.Folder = oMoveTarget
End With

'Update the server and display progress dialog
colRules.Save
oRule.Execute ShowProgress:=True
End Sub
