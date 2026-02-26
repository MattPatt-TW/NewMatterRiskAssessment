<?xml version="1.0" encoding="UTF-8"?>
<tfb>
  <RiskMatterV2>
    <Init>
      <![CDATA[
import clr
import System

clr.AddReference("System")            # for new MRA Edit Template tab code
clr.AddReference("WindowsBase")       # for new MRA Edit Template tab code

clr.AddReference('mscorlib')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('System.Windows.Forms')

from System import DateTime, Action, Convert, DBNull
from System.Diagnostics import Process
from System.Globalization import DateTimeStyles
from System.Collections.Generic import Dictionary
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import Controls, Forms, LogicalTreeHelper, RoutedEventHandler
from System.Windows import Data, UIElement, Visibility, Window, GridLength, GridUnitType
from System.Windows.Controls import Button, Canvas, GridView, GridViewColumn, ListView, Orientation, DataGridCellInfo
from System.Windows.Data import Binding, CollectionView, ListCollectionView, PropertyGroupDescription, CollectionViewSource
from System.Windows.Forms import SelectionMode, MessageBox, MessageBoxButtons, DialogResult, MessageBoxIcon
from System.Windows.Input import KeyEventHandler
from System.Windows.Media import Brush, Brushes
from System.Windows.Threading import DispatcherPriority

# Global Variables
UserIsHod = False
UserSelfApproves = False
UserIsAnApprovalUser = False
UserCanReviewFiles = False
RiskAndITUsers = ['MP', 'LD1', 'AH1', 'JM1', 'AF1']

# for newer structure of MRA
MRA_ANSWERS_BY_QID = {}   # temp to store list of 'MRA_PREVIEW_ANSWER_ROW' dicts for current TemplateID being previewed
MRA_QUESTIONS_LIST = []   # temp to store list of 'MRA_PREVIEW_QUESTION_ROW' dicts for current TemplateID being previewed
_preview_combo_syncing = False   # temp variable to avoid triggering 'SelectionChanged' event when we're programmatically changing combo box selections in the MRA preview screen

# dict to hold current matter answer/comment by QuestionID
MRA_MATTER_SELECTIONS_BY_QID = {}   # qid -> {"AnswerID": int|None, "Comments": str}

UNSELECTED = -1

# Standardised MRA Outcomes
OUTCOME_AUTO_APPROVE = "AUTO_APPROVE"
OUTCOME_REQUEST_HOD  = "REQUEST_HOD"
OUTCOME_SUBMIT_STD   = "SUBMIT_STD"
OUTCOME_ON_BEHALF    = "SUBMIT_ON_BEHALF"

# Global Variables for session tokens
token1 = 0

# # # #   O N   L O A D   E V E N T   # # # # -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 
def myOnLoadEvent(s, event):
  
  # Update UserIsHod
  # Returns True if user can approve the fee earner of this matter
  global UserIsHod
  global UserSelfApproves
  global UserIsAnApprovalUser
  global UserCanReviewFiles
  UserIsHod = canUserApproveFeeEarner(UserToCheck = _tikitUser, FeeEarner = tb_FERef.Text)
  UserSelfApproves = canApproveSelf(userToCheck = _tikitUser)
  UserIsAnApprovalUser = isUserAnApprovalUser(userToCheck = _tikitUser)
  UserCanReviewFiles = canUserReviewFiles(userToCheck = _tikitUser)

  # refresh main overview DataGrids
  dg_MRAFR_Refresh(s, event)
  dgCA_Overview_Refresh(s, event)
  
  # Hide MRA and FE tabs (user needs to select from 'Overview' list, and click 'Edit' or 'View' to then show details
  ti_MRA.Visibility = Visibility.Collapsed
  ti_FR.Visibility = Visibility.Collapsed
  
  # put current user details into fields on Corrective Actions area
  tb_CurrUser.Text = _tikitUser
  tb_CurrUserName.Text = _tikitResolver.Resolve("[SQL: SELECT FullName FROM Users WHERE Code = '{0}']".format(_tikitUser))
  # NB: Tag="SQL: SELECT '[curentuser.code]'" didn't work on XAML (hence above)
  # Tag="SQL: SELECT '[currentuser.fullname]'"
  
  # set current risk status
  setMasterRiskStatus(s, event)
  
  # as only IT or Risk should be able to delete, hide the 'delete' button for everyone else:
  if _tikitUser not in RiskAndITUsers:
    sep_Delete.Visibility = Visibility.Collapsed
    btn_DeleteSelected_MRAFR.Visibility = Visibility.Collapsed

  # wire up new 'New' button popup
  icTemplates.AddHandler(Button.ClickEvent,
        RoutedEventHandler(TemplateButton_Click))

  showHide_ApproveMRAbutton(s, event)
  POPULATE_AGENDA_NAMES(_tikitSender, 'onLoad')
  refresh_CaseDocs(_tikitSender, '')
  return

# # # #   END:   O N   L O A D   E V E N T   # # # # -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --

# # # # # # # # # # # #  C A S E   D O C S   -   F U N C T I O N S  # # # # # # # # # # # # # # 

def opt_EntityOrMatterDocs_Clicked(s, event):
  # Linked to XAML control: opt_CaseDocs_Entity
  # needs to re-populate 'Agenda' combo box to only show Entity level docs
  POPULATE_AGENDA_NAMES(s, 'onLoad')
  refresh_CaseDocs(s, event)
  return

def CaseDoc_SelectionChanged(s, event):
  if dg_CaseManagerDocs.SelectedIndex == -1:
    btn_OpenCaseDoc.IsEnabled = False
  else:
    if dg_CaseManagerDocs.SelectedItem['Path'] != '':
      btn_OpenCaseDoc.IsEnabled = True
  return
  

def open_Selected_CaseDoc(s, event):
  #MessageBox.Show('Testing open Case Doc button')
  tmpPath = dg_CaseManagerDocs.SelectedItem['Path']
  tmpName = dg_CaseManagerDocs.SelectedItem['Desc']
  
  #MessageBox.Show('Testing open Case Doc button. \nName:' + tmpName + '\nPath:' + tmpPath)
  
  if tmpPath == '':
    MessageBox.Show("There doesn't appear to be a path to this document: \n" + str(tmpName))
  else:
    System.Diagnostics.Process.Start(r'{}'.format(tmpPath))
  return
  

# Case Docs DataGrid
class CaseDocs(object):
  def __init__(self, mySID, mySDesc, mySCreated, mySPath):
    self.sID = mySID
    self.sDescription = mySDesc
    self.sCreated = mySCreated
    self.sDocPath = mySPath
    return

  def __getitem__(self, index):
    if index == 'ID':
      return self.sID
    elif index == 'Desc':
      return self.sDescription
    elif index == 'Created':
      return self.sCreated 
    elif index == 'Path':
      return self.sDocPath 


def refresh_CaseDocs(s, event):
  
  # if there's nothing selected, exit function
  if cbo_AgendaName.SelectedIndex == -1:
    return
  
  # otherwise, get the selected item ID and continue populating...
  tmpAgendaName = cbo_AgendaName.SelectedItem['ID']
  
  # Get the SQL
  sSQL = """SELECT Cm_CaseItems.ItemID, Cm_CaseItems.Description, Cm_CaseItems.CreationDate,  Cm_Steps.FileName 
            FROM Cm_CaseItems 
              INNER JOIN Cm_Steps ON Cm_Steps.ItemID = Cm_CaseItems.ItemID 
            WHERE ParentID = {0} AND Cm_Steps.FileName <> '' ORDER BY Cm_CaseItems.ItemOrder """.format(tmpAgendaName)
  sItem = []
  
  # Open and store items in code
  _tikitDbAccess.Open(sSQL)
  
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        aID = 0 if dr.IsDBNull(0) else dr.GetValue(0)
        aDesc = '' if dr.IsDBNull(1) else dr.GetString(1)
        aDate = 0 if dr.IsDBNull(2) else dr.GetValue(2)
        aPath = '' if dr.IsDBNull(3) else dr.GetString(3)
      
        sItem.append(CaseDocs(aID, aDesc, aDate, aPath))
    
    dr.Close()
  _tikitDbAccess.Close

  # Set 'Source' and close db connection
  dg_CaseManagerDocs.ItemsSource = sItem
  return
  

# Agenda Items
class cboAgendaNames(object):
  def __init__(self, myAgendaID, myAgendaName, myDefault):
    self.AgendaID = myAgendaID
    self.AgendaName = myAgendaName
    self.mIsDefault = myDefault

    if myAgendaName == 'Case History':
      self.mIsDefault = 1
    else:
      self.mIsDefault = myDefault
    return

  def __getitem__(self, index):
    if index == 'ID':
      return self.AgendaID
    elif index == 'Name':
      return self.AgendaName
    elif index == 'Default':
      return self.mIsDefault


def POPULATE_AGENDA_NAMES(s, event):
  # This function populates the combo box housing the 'Agendas' for the current matter
  # Updated 31st Oct 2024: Including ability to look at 'Entity' level docs too via 2 new Radio buttons above the 'Agenda' combo box

  # if the option for Entity is selected
  if opt_CaseDocs_Entity.IsChecked == True:
    # set matter number to zero (for entity-level docs)
    tmpMatterNo = 0
  else:
    # assumes 'matter' option selected, so set to active matter
    tmpMatterNo = _tikitMatter

  # form SQL 
  mySQL = """SELECT Cm_CaseItems.Description, Cm_Agendas.ItemID, Cm_Agendas.Default_Agenda 
            FROM Cm_Agendas LEFT JOIN Cm_CaseItems ON Cm_Agendas.ItemID = Cm_CaseItems.ItemID 
            WHERE Cm_Agendas.EntityRef = '{0}' AND Cm_Agendas.MatterNo = {1} 
            ORDER BY Cm_CaseItems.Description""".format(_tikitEntity, tmpMatterNo)
  
  # open SQL and create a new list object to hold items
  _tikitDbAccess.Open(mySQL)
  itemA = []
  
  # iterate over returned rows from SQL putting values into temp variables
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          iAgendaName = '-' if dr.IsDBNull(0) else dr.GetString(0)
          iAgendaID = '-' if dr.IsDBNull(1) else dr.GetValue(1)
          iDefault = 0 if dr.IsDBNull(2) else dr.GetValue(2)
          # finally add item to our temp list
          itemA.append(cboAgendaNames(iAgendaID, iAgendaName, iDefault))
    dr.Close()
  _tikitDbAccess.Close()

  # Set set source of the Agenda Names combo box to the list of items we created
  cbo_AgendaName.ItemsSource = itemA

  if event == 'onLoad':
    #MessageBox.Show("Onload Test - this works, so need to add code to auto select default")
    tmpCount = -1
    
    for x in cbo_AgendaName.Items:
      tmpCount += 1
      if x['Default'] == 1:
        #MessageBox.Show("This one is the 'default': " + str(x['Name']))
        cbo_AgendaName.SelectedIndex = tmpCount
  return
        
# #  END: C A S E   D O C S   -   F U N C T I O N S  # #


def showHide_ApproveMRAbutton(s, event):
  # If no row is selected in the data grid, disable the buttons and exit
  if dg_MRAFR.SelectedIndex == -1:
    btn_HOD_Approval_MRA.IsEnabled = False
    btn_HOD_Approval_MRA1.IsEnabled = False
    return

  # Retrieve necessary fields from the selected item
  tmpSelectedRiskRating = dg_MRAFR.SelectedItem['RiskRating']
  tmpSelectedStatus = dg_MRAFR.SelectedItem['Status']
  tmpSelectedApprovedByHOD = dg_MRAFR.SelectedItem['ApprovedByHod']
  tmpSelectedType = dg_MRAFR.SelectedItem['Type']

  # Check if the selected item is of type 'Matter Risk Assessment'
  if 'Matter Risk Assessment' not in tmpSelectedType:
    btn_HOD_Approval_MRA.IsEnabled = False
    btn_HOD_Approval_MRA1.IsEnabled = False
    return

  if (tmpSelectedRiskRating == 'High' and 
      tmpSelectedStatus == 'Complete' and 
      tmpSelectedApprovedByHOD == 'No'):
        
    currentUser = _tikitUser
    isApprovalUser = isUserAnApprovalUser(userToCheck=currentUser)
        
    # Check if the current user can approve the fee earner
    if canUserApproveFeeEarner(UserToCheck=currentUser, FeeEarner=tb_FERef.Text):
      btn_HOD_Approval_MRA.IsEnabled = isApprovalUser
      btn_HOD_Approval_MRA1.IsEnabled = isApprovalUser
      return

  # If any condition is not met, disable the buttons
  btn_HOD_Approval_MRA.IsEnabled = False
  btn_HOD_Approval_MRA1.IsEnabled = False
  return


def setMasterRiskStatus(s, event):
  # This function sets the main label on the Overview sheet according to current matter Risk Status
  
  tmpSQL = "SELECT CASE RiskOpening WHEN 1 THEN 'Low' WHEN 2 THEN 'Medium' WHEN 3 THEN 'High' ELSE 'NotSet' END FROM Matters WHERE EntityRef = '{0}' AND Number = {1}".format(_tikitEntity, _tikitMatter) 
  lbl_OV_RiskStatus.Content = runSQL(tmpSQL, False, '', '')
  
  # display or hide the 'advisory' text for High Risk matters as appropriate
  #  (this just advises that for HighRisk matters, a new one is req each month)
  if lbl_OV_RiskStatus.Content == 'High':
    lbl_RiskScore_AdvisoryText.Visibility = Visibility.Visible
  else:
    lbl_RiskScore_AdvisoryText.Visibility = Visibility.Collapsed
  return


def MRA_setStatus(idToUpdate, newStatus):
  # This function will set the status of the active MRA accordingly
  
  if int(idToUpdate) > 0 and len(newStatus) > 0:
    mySQL = "UPDATE Usr_MRAv2_MatterHeader SET Status = '{0}', SubmittedBy = '{1}', SubmittedDate = GETDATE() WHERE mraID = {2}".format(newStatus, _tikitUser, idToUpdate)
    runSQL(mySQL, True, "There was an error updating the Status for this Matter Risk Assessment", "Error: MRA_setStatus")
    lbl_MRA_Status.Content = newStatus
  return
  

# # # #   O V E R V I E W    TAB   # # # # 

class MRAFR(object):
  def __init__(self, mymraID, myTemplateID, myName, myExpiryDate, myScore, myRiskR, 
               myAppByHOD, myQCount, myQOS, myStatus, mySubbedBy, mySubbedOn, 
               myScoreTriggerMed, myScoreTriggerHigh, myFRReviewer, myType):
    self.mraID = mymraID
    self.Name = myName
    self.Status = myStatus
    self.ExpiryDate = myExpiryDate
    if myRiskR == 0:
      self.RiskRating = "-Not Set-"
    elif myRiskR == 1:
      self.RiskRating = "Low"
    elif myRiskR == 2:
      self.RiskRating = "Medium"
    elif myRiskR == 3:
      self.RiskRating = "High"
    
    if myAppByHOD == 'Y':
      self.AppByHOD = 'Yes'
    else:
      self.AppByHOD = 'No'

    self.QCount = myQCount
    self.QOutstanding = myQOS
    self.FRReviewer = myFRReviewer
    self.SubbedBy = mySubbedBy
    self.SubbedOn = mySubbedOn
    self.ScoreTriggerMed = myScoreTriggerMed
    self.ScoreTriggerHigh = myScoreTriggerHigh
    self.TemplateID = myTemplateID
    self.Score = myScore
    self.Type = myType
    return
    
  def __getitem__(self, index):
    if index == 'mraID':
      return self.mraID
    elif index == 'Name':
      return self.Name
    elif index == 'TemplateID':
      return self.TemplateID
    elif index == 'Status':
      return self.Status
    elif index == 'Expiry':
      return self.ExpiryDate
    elif index == 'QCount':
      return self.QCount
    elif index == 'QOS':
      return self.QOutstanding
    elif index == 'FR Reviewer':
      return self.FRReviewer
    elif index == 'RiskRating':
      return self.RiskRating
    elif index == 'SubbedBy':
      return self.SubbedBy
    elif index == 'SubbedOn':
      return self.SubbedOn
    elif index == 'ApprovedByHod':
      return self.AppByHOD
    elif index == 'ScoreTriggerMedium':
      return self.ScoreTriggerMed
    elif index == 'ScoreTriggerHigh':
      return self.ScoreTriggerHigh
    elif index == 'Type':
      return self.Type
    elif index == 'Score':
      return self.Score
    else:
      return ''

def dg_MRAFR_Refresh(s, event):
  # This funtion populates the main Matter Risk Assessment & File Review data grid 

  # SQL to populate datagrid
  getTableSQL = """WITH
                  /* =========================
                    MRAv2 branch
                    ========================= */
                  MatterHeader AS (
                      SELECT MH.mraID, MH.Name, MH.Status, MH.ApprovedByHOD, MH.DateAdded,
                            MH.ExpiryDate, MH.SubmittedBy, MH.SubmittedDate, MH.TemplateID
                      FROM Usr_MRAv2_MatterHeader MH  
                      WHERE MH.EntityRef = '{entRef}' AND MH.MatterNo = {matNo}
                  ),
                  AnsweredQuestions AS (
                      SELECT MD.mraID, MD.QuestionID, MD.AnswerID, MD.Score
                      FROM Usr_MRAv2_MatterDetails MD
                      WHERE MD.EntityRef = '{entRef}' AND MD.MatterNo = {matNo}
                  ),
                  TemplateStructure AS (
                      SELECT T.TemplateID, T.QuestionID
                      FROM Usr_MRAv2_Templates T
                      INNER JOIN MatterHeader MH ON T.TemplateID = MH.TemplateID
                  ),
                  AnswerAgg AS (
                      SELECT mraID,
                            SUM(ISNULL(Score,0)) AS Score,
                            COUNT(*)             AS AnswerCount,
                            COUNT(DISTINCT QuestionID) AS AnsweredQuestionCount
                      FROM AnsweredQuestions
                      GROUP BY mraID
                  ),
                  TemplateAgg AS (
                      SELECT TemplateID,
                            COUNT(DISTINCT QuestionID) AS QCount
                      FROM TemplateStructure
                      GROUP BY TemplateID
                  ),
                  TemplateHeader AS (
                      SELECT TD.TemplateID,
                            MAX(TD.ScoreMediumTrigger) AS ScoreTriggerMed,
                            MAX(TD.ScoreHighTrigger)   AS ScoreTriggerHigh
                      FROM Usr_MRAv2_TemplateDetails TD
                      GROUP BY TD.TemplateID
                  ),

                  MRAv2_Result AS (
                      SELECT
                          '00-mraID'            = MH.mraID,
                          '01-TemplateID'       = MH.TemplateID,
                          '02-Name'             = MH.Name,
                          '03-ExpiryDate'       = MH.ExpiryDate,
                          '04-Score'            = ISNULL(AA.Score, 0),
                          '05-RiskRating'       = CASE WHEN ISNULL(AA.Score, 0) < TH.ScoreTriggerMed THEN 1
                                                      WHEN ISNULL(AA.Score, 0) >= TH.ScoreTriggerHigh THEN 3
                                                      ELSE 2 END,
                          '06-ApprovedByHOD'    = MH.ApprovedByHOD,
                          '07-QCount'           = TA.QCount,
                          '08-OS Count'         = TA.QCount - ISNULL(AA.AnswerCount, 0),
                          '09-Status'           = MH.Status,
                          '10-SubmittedBy'      = ISNULL(CONCAT('(', U.Code, ') ', U.FullName), ''),
                          '11-SubmittedDate'    = MH.SubmittedDate,
                          '12-ScoreTriggerMed'  = TH.ScoreTriggerMed,
                          '13-ScoreTriggerHigh' = TH.ScoreTriggerHigh,
                          '14-DateAdded'        = MH.DateAdded,
                          '15-FR Reviewer'      = CAST(NULL AS varchar(100)),  
                          '16-Type'             = 'Matter Risk Assessment'
                      FROM MatterHeader MH
                      LEFT JOIN AnswerAgg      AA ON AA.mraID      = MH.mraID
                      LEFT JOIN TemplateAgg    TA ON TA.TemplateID = MH.TemplateID
                      LEFT JOIN TemplateHeader TH ON TH.TemplateID = MH.TemplateID
                      LEFT OUTER JOIN Users     U ON MH.SubmittedBy = U.Code
                  ),

                  /* =========================
                    File Review branch
                    ========================= */
                  FR_Header AS (
                      SELECT MRAO.ID, MRAO.TypeID, TT.TypeName,
                          MRAO.ExpiryDate, MRAO.Score, MRAO.RiskRating,
                          MRAO.ApprovedByHOD, MRAO.Status, MRAO.SubmittedBy,
                          MRAO.SubmittedDate, MRAO.DateAdded, MRAO.FR_Reviewer,
                          MRAO.EntityRef, MRAO.MatterNo
                      FROM Usr_MRA_Overview MRAO
                      LEFT JOIN Usr_MRA_TemplateTypes TT ON MRAO.TypeID = TT.TypeID
                      WHERE MRAO.EntityRef = '{entRef}' AND MRAO.MatterNo = {matNo}
                        AND TT.Is_MRA = 'N'
                  ),
                  FR_DetailAgg AS (
                      SELECT D.OV_ID, QCount   = COUNT(*),
                          OSCount  = SUM(CASE WHEN D.SelectedAnswerID = -1 THEN 1 ELSE 0 END)
                      FROM Usr_MRA_Detail D
                      WHERE D.EntityRef = '{entRef}' AND D.MatterNo = {matNo}
                      GROUP BY D.OV_ID
                  ),
                  FR_Result AS (
                      SELECT '00-mraID'         = H.ID,
                          '01-TemplateID'       = H.TypeID,
                          '02-Name'             = H.TypeName,
                          '03-ExpiryDate'       = H.ExpiryDate,
                          '04-Score'            = H.Score,
                          '05-RiskRating'       = H.RiskRating,
                          '06-ApprovedByHOD'    = H.ApprovedByHOD,
                          '07-QCount'           = ISNULL(A.QCount, 0),
                          '08-OS Count'         = ISNULL(A.OSCount, 0),
                          '09-Status'           = ISNULL(H.Status, 'Draft'),
                          '10-SubmittedBy'      = ISNULL(CONCAT('(', U.Code, ') ', U.FullName), ''),
                          '11-SubmittedDate'    = H.SubmittedDate,
                          '12-ScoreTriggerMed'  = 0,
                          '13-ScoreTriggerHigh' = 0,
                          '14-DateAdded'        = H.DateAdded,
                          '15-FR Reviewer'      = ISNULL(CONCAT('(', UR.Code, ') ', UR.FullName), 'N/A'), 
                          '16-Type'             = 'File Review'
                      FROM FR_Header H
                      LEFT JOIN FR_DetailAgg A ON A.OV_ID = H.ID
                      LEFT JOIN Users U        ON H.SubmittedBy  = U.Code
                      LEFT JOIN Users UR       ON H.FR_Reviewer  = UR.Code
                  )
                  SELECT *
                  FROM MRAv2_Result
                  UNION ALL
                  SELECT *
                  FROM FR_Result
                  ORDER BY [14-DateAdded] DESC;""".format(entRef=_tikitEntity, matNo=_tikitMatter)
  
  tmpText = "'Matter Risk Assessment(s)' or 'File Review(s)'"
    
  #MessageBox.Show("GetTableSQL:\n" + getTableSQL, "Debug: Populating Matter Risk Assessment and File Review")
  
  tmpItem = []
  
  _tikitDbAccess.Open(getTableSQL)
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          imraID = 0 if dr.IsDBNull(0) else dr.GetValue(0)
          iTemplateID = 0 if dr.IsDBNull(1) else dr.GetValue(1)
          iName = '' if dr.IsDBNull(2) else dr.GetString(2)
          iExpiry = 0 if dr.IsDBNull(3) else dr.GetValue(3)
          iScore = 0 if dr.IsDBNull(4) else dr.GetValue(4) 
          iRiskR = 0 if dr.IsDBNull(5) else dr.GetValue(5)
          iAppByHOD = '' if dr.IsDBNull(6) else dr.GetString(6)
          iQCount = 0 if dr.IsDBNull(7) else dr.GetValue(7) 
          iOSQs = 0 if dr.IsDBNull(8) else dr.GetValue(8) 
          iStatus = '' if dr.IsDBNull(9) else dr.GetString(9)
          iSubBy = '' if dr.IsDBNull(10) else dr.GetString(10)
          iSubOn = 0 if dr.IsDBNull(11) else dr.GetValue(11)
          iScoreTrigMed = 0 if dr.IsDBNull(12) else dr.GetValue(12)
          iScoreTrigHigh = 0 if dr.IsDBNull(13) else dr.GetValue(13)
          iDateAdded = 0 if dr.IsDBNull(14) else dr.GetValue(14)
          iFRR = '' if dr.IsDBNull(15) else dr.GetString(15)
          iType = '' if dr.IsDBNull(16) else dr.GetString(16)

          tmpItem.append(MRAFR(mymraID=imraID, myTemplateID=iTemplateID, myName=iName, myExpiryDate=iExpiry, myScore=iScore, 
                               myRiskR=iRiskR, myAppByHOD=iAppByHOD, myQCount=iQCount, myQOS=iOSQs, 
                               myStatus=iStatus, mySubbedBy=iSubBy, mySubbedOn=iSubOn, 
                               myScoreTriggerMed=iScoreTrigMed, myScoreTriggerHigh=iScoreTrigHigh, myFRReviewer=iFRR, myType=iType))
    dr.Close()
  
  #close db connection
  _tikitDbAccess.Close()
  dg_MRAFR.ItemsSource = tmpItem
  
  if dg_MRAFR.Items.Count > 0:
    dg_MRAFR.Visibility = Visibility.Visible
    tb_NoMRAFR.Text = ""
    tb_NoMRAFR.Visibility = Visibility.Hidden
    btn_CopySelected_MRAFR.IsEnabled = True
    btn_View_MRAFR.IsEnabled = True
    btn_Edit_MRAFR.IsEnabled = True
    btn_DeleteSelected_MRAFR.IsEnabled = True
  else:
    tb_NoMRAFR.Text = "No {0}'s currently exist on this matter - please click the '+ New' button to create new".format(tmpText)
    tb_NoMRAFR.Visibility = Visibility.Visible
    dg_MRAFR.Visibility = Visibility.Hidden
    btn_CopySelected_MRAFR.IsEnabled = False
    btn_View_MRAFR.IsEnabled = False
    btn_Edit_MRAFR.IsEnabled = False
    btn_DeleteSelected_MRAFR.IsEnabled = False
  return


def dg_MRAFR_Refresh_and_Reselect(mraIDToReselect=None):
  # This function will refresh the MRA/FR datagrid but attempt to keep the same item selected (if it still exists)

  tmpSelectedID = mraIDToReselect
  if tmpSelectedID is None and dg_MRAFR.SelectedItem is not None:
    tmpSelectedID = getattr(dg_MRAFR.SelectedItem, 'mraID', None) or dg_MRAFR.SelectedItem['mraID']
  
  dg_MRAFR_Refresh(None, None)

  if tmpSelectedID is None:
    return

  def do_select():
    # ensure containers exist
    dg_MRAFR.UpdateLayout()
    dg_MRAFR.Focus()

    dg_MRAFR.SelectedValue = tmpSelectedID
    dg_MRAFR.UpdateLayout()

    if dg_MRAFR.SelectedItem is not None:
      dg_MRAFR.ScrollIntoView(dg_MRAFR.SelectedItem)
  
  dg_MRAFR.Dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle, Action(do_select))
  return


def dg_MRAFR_SelectionChanged(s, event):
  # This function will populate the label controls to temp store ID and Name

  global UserIsHod

  if dg_MRAFR.SelectedIndex > -1:
    lbl_MRAFR_ID.Content = dg_MRAFR.SelectedItem['mraID']
    lbl_MRAFR_Name.Content = dg_MRAFR.SelectedItem['Name']
    #refresh_MRA_LMH_ScoreThresholds(s, event)
    tmpType = str(dg_MRAFR.SelectedItem['Type'])
    tmpSearchText = str('Matter Risk Assessment')

    if tmpSearchText in tmpType:
      #MessageBox.Show("Type is {0}".format(tmpType), "DEBUGGING")
      # only enable the 'Duplicate Selected' button for NMRA's (not applicable to FR's)    
      btn_CopySelected_MRAFR.IsEnabled = True 
    else:
      btn_CopySelected_MRAFR.IsEnabled = False

    if dg_MRAFR.SelectedItem['Status'] == 'Complete':
      btn_Edit_MRAFR.IsEnabled = False
    elif dg_MRAFR.SelectedItem['Status'] == 'With FE':
      #MessageBox.Show(str(UserIsHod))
      btn_Edit_MRAFR.IsEnabled = UserIsHod
    else:
      btn_Edit_MRAFR.IsEnabled = True
  else:
    lbl_MRAFR_ID.Content = ''
    lbl_MRAFR_Name.Content = ''
    btn_Edit_MRAFR.IsEnabled = False
    btn_CopySelected_MRAFR.IsEnabled = False
    
  showHide_ApproveMRAbutton(s, event)
  return

class TemplateOption(object):
  def __init__(self, template_id, name, validity_days, iType, iQCount):
      self.TemplateID = template_id
      self.Name = name
      self.ValidityDays = validity_days
      self.Type = iType
      self.QCount = iQCount

def load_templates_for_case_type(matterCaseType):
  sql = """SELECT NewMRA.ID, NewMRA.Name, NewMRA.iLockDays, NewMRA.Type, NewMRA.QCount FROM (
              SELECT 'ID' = TD.TemplateID, 'Name' = TD.Name, 
                  'iLockDays' = TD.DaysUntil_IncompleteLock, 
                  'Type' = 'MRA', 
                  'QCount' = (SELECT COUNT(QuestionID) FROM Usr_MRAv2_Templates WHERE TemplateID = TD.TemplateID)
              FROM Usr_MRAv2_TemplateDetails TD
                JOIN Usr_MRAv2_CaseTypeDefaults CTD ON TD.TemplateID = CTD.TemplateID
              WHERE CTD.CaseTypesCode = {CaseType}
              UNION ALL
              SELECT 'ID' = TT.TypeID, 'Name' = TT.TypeName, 
                  'iLockDays' = 60, 'Type' = 'FileReview', 
                  'QCount' = (SELECT COUNT(ID) FROM Usr_MRA_TemplateQs WHERE TypeID = TT.TypeID)
              FROM Usr_MRA_TemplateTypes TT 
                JOIN Usr_MRA_CaseType_Defaults CTD ON TT.TypeID = CTD.TemplateID
              WHERE TT.Is_MRA = 'N' AND CTD.TypeName = 'File Review' AND CTD.CaseTypeID = {CaseType}
              ) AS NewMRA""".format(CaseType=matterCaseType)

  items = ObservableCollection[object]()
  # standard P4W 'execute' SQL pattern
  _tikitDbAccess.Open(sql)
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        template_id = 0 if dr.IsDBNull(0) else dr.GetValue(0)
        name = '' if dr.IsDBNull(1) else dr.GetString(1)
        validity_days = 0 if dr.IsDBNull(2) else dr.GetValue(2)
        iType = '' if dr.IsDBNull(3) else dr.GetString(3)
        iQCount = 0 if dr.IsDBNull(4) else dr.GetValue(4)
        items.Add(TemplateOption(template_id, name, validity_days, iType, iQCount))
    dr.Close()
  _tikitDbAccess.Close()

  return items

def btnNew_Click(sender, args):
  # ToggleButton behaviour: use IsChecked to decide open state
  is_open = bool(sender.IsChecked)

  if is_open:
    matterCaseType = tb_CaseTypeRef.Text
    templates = load_templates_for_case_type(matterCaseType)

    # Bind list
    icTemplates.ItemsSource = templates

    # Open popup
    popTemplates.IsOpen = True
  else:
    popTemplates.IsOpen = False

def TemplateButton_Click(sender, args):
  # Because we used AddHandler on icTemplates, sender may be the ItemsControl.
  # args.OriginalSource should be the actual Button (or something inside it).
  global UserCanReviewFiles

  btn = args.OriginalSource
  # Sometimes OriginalSource is a TextBlock inside the Button; walk up to Button
  while btn is not None and not isinstance(btn, Button):
    btn = getattr(btn, "TemplatedParent", None) or getattr(btn, "Parent", None)

  if btn is None:
    return

  opt = btn.Tag  # <-- TemplateOption
  template_id = opt.TemplateID
  tName = opt.Name
  validity_days = opt.ValidityDays
  iType = opt.Type
  iQCount = opt.QCount

  # if no questions on this template, alert user and exit
  if iQCount == 0:
    MessageBox.Show("The template '{0}' doesn't have any questions associated with it, so you can't create a new {1} using this template.\n\nPlease contact your system administrator to resolve this.".format(tName, iType), "Error: No Questions on Template")
    return

  # Choose behaviour based on type
  if iType == "MRA":
    AddNew_MRA(templateID=template_id, templateName=tName, templateValidityDays=validity_days)
    #MessageBox.Show("You clicked to add a new MRA ({0}) with TemplateID: {1}".format(tName, template_id), "DEBUG: Add new MRA")
    
    #MRA_AddNew(template_id, validity_days)  # whatever signature you need
  elif iType == "FileReview":
    if UserCanReviewFiles:
      AddNew_FileReview(templateID=template_id, templateName=tName, templateValidityDays=validity_days)
    else:
      MessageBox.Show("Your user account doesn't have permissions to create new File Reviews. Please contact your system administrator.", "Permission Denied: Create File Review")

  # Close popup
  popTemplates.IsOpen = False

def CancelPopup_Click(sender, args):
  popTemplates.IsOpen = False

def popTemplates_Closed(sender, args):
  # Ensure the toggle button pops back up
  if btnNew is not None:
    btnNew.IsChecked = False

# - END OF 'ADD NEW' Button controls #
def AddNew_MRA(templateID, templateName, templateValidityDays):
  # This function will add a new row to the 'Matter Risk Assessment' data grid

  sql = """WITH MainTemplates AS (
                  SELECT TD.TemplateID, TD.Name
                  FROM Usr_MRAv2_TemplateDetails TD
                  WHERE TD.TemplateID = {templateID}
                  ), 
                ExistingMatterRows AS (
                  SELECT MH.TemplateID, MH.mraID 
                  FROM Usr_MRAv2_MatterHeader MH
                  WHERE MH.EntityRef = '{entRef}' AND MH.MatterNo = {matNo}
                  )
            INSERT INTO Usr_MRAv2_MatterHeader (EntityRef, MatterNo, mraID, Name, Status, ApprovedByHOD, DateAdded, TemplateID, ExpiryDate)
            OUTPUT inserted.mraID 
            SELECT '{entRef}', {matNo}, (SELECT ISNULL(MAX(mraID), 0) + 1 FROM Usr_MRAv2_MatterHeader), 
                CONCAT(MT.Name, ' (', CONVERT(NVARCHAR, COUNT(EMR.mraID) + 1), ')'), 
                'Draft', 'N', GETDATE(), {templateID}, DATEADD(DAY, {validityDays}, GETDATE())
            FROM MainTemplates MT
                LEFT OUTER JOIN ExistingMatterRows EMR ON MT.TemplateID = EMR.TemplateID
            GROUP BY MT.Name""".format(entRef=_tikitEntity, matNo=_tikitMatter, templateID=templateID, validityDays=templateValidityDays)

  #try:
  newMRAid = runSQL(codeToRun=sql, useAlternativeResolver=True, 
                    errorMsgText="There was an error adding new item to overview table", 
                    errorMsgTitle="Error: Add New MRA...", showError=True, returnType='Int')
  #except Exception as e:
  #  MessageBox.Show("There was an error adding new item to overview table, using SQL:\n{0}\n\nError details:\n{1}".format(sql, str(e)), "Error: Add New MRA...")
  #  return
  MessageBox.Show("SQL to add new MRA:\n{0}\n\nReturned new MRA ID: {1}".format(sql, newMRAid), "DEBUG: Add New MRA - SQL and Result")

  if int(newMRAid) > 0:
    dg_MRAFR_Refresh_and_Reselect(mraIDToReselect=newMRAid)
  #MessageBox.Show("You clicked to add a new MRA with TemplateID: {0}".format(templateID), "DEBUG: Add new MRA")
  return
  
def AddNew_FileReview(templateID, templateName, templateValidityDays):
  # This function will add a new row to the 'Matter Risk Assessment / File Review' data grid
  #! In porocess of being re-written / updated (whilst we're not changing any underlying code on 
  #!  file review process, we can afford to smarten up some of this code as it's looking a bit jank
  #!  (with far too many SQL calls!))
  
  newHeaderID_sql = """WITH MainTemplates AS (
                            SELECT TT.TypeID, TT.TypeName
                            FROM Usr_MRA_TemplateTypes TT
                            WHERE TT.TypeID = {templateID}
                            ), 
                          ExistingMatterRows AS (
                            SELECT MRAO.TypeID, MRAO.ID 
                            FROM Usr_MRA_Overview MRAO
                            WHERE MRAO.EntityRef = '{entRef}' AND MRAO.MatterNo = {matNo}
                            )
                        INSERT INTO Usr_MRA_Overview (EntityRef, MatterNo, TypeID, ExpiryDate, LocalName, Score, RiskRating, ApprovedByHOD, DateAdded, FR_Reviewer, Status)
                         OUTPUT inserted.ID
                        SELECT '{entRef}', {matNo}, {templateID}, DATEADD(DAY, {validityDays}, GETDATE()), 
                          CONCAT(MT.TypeName, ' (', CONVERT(NVARCHAR, COUNT(EMR.ID) + 1), ')'),
                          0, 0, 'N', GETDATE(), '{cUser}', 'Draft'
                        FROM MainTemplates MT
                          LEFT OUTER JOIN ExistingMatterRows EMR ON MT.TypeID = EMR.TypeID
                        GROUP BY MT.TypeName""".format(templateID=templateID, entRef=_tikitEntity, 
                                                       matNo=_tikitMatter, validityDays=templateValidityDays, cUser=_tikitUser)

  #try:
  newHeaderID = runSQL(codeToRun=newHeaderID_sql, useAlternativeResolver=True,
                       errorMsgText="There was an error adding new item to overview table", 
                       errorMsgTitle="Error: Add New File Review...", showError=True, returnType='Int')
  #except Exception as e:
  #  MessageBox.Show("There was an error adding new item to overview table, using SQL:\n{0}\n\nError details:\n{1}".format(newHeaderID_sql, str(e)), "Error: Add New File Review...")
  #  return

  if int(newHeaderID) < 0:
    MessageBox.Show("There was an error adding new item to overview table, using SQL:\n{0}".format(newHeaderID_sql), "Error: Add New File Review...")
    return
  
  # finally add the Questions (to the Details table)
  insert_Qs_SQL = """INSERT INTO Usr_MRA_Detail (EntityRef, MatterNo, OV_ID, QuestionID, AnswerListToUse, SelectedAnswerID, CurrentAnswerScore, DisplayOrder)  
                     SELECT '{0}', {1}, {2}, ID, AnswerList, -1, 0, DisplayOrder FROM Usr_MRA_TemplateQs 
                     WHERE TypeID = {3}""".format(_tikitEntity, _tikitMatter, newHeaderID, templateID)
  try:
    _tikitResolver.Resolve("[SQL: {0}]".format(insert_Qs_SQL))
  except:
    MessageBox.Show("There was an error copying the Questions to the 'Details' table, using SQL:\n{0}".format(insert_Qs_SQL), "Error: Add New File Review...")
    return  
    
  # refresh data grid 
  dg_MRAFR_Refresh_and_Reselect(mraIDToReselect=newHeaderID)
  return



def btn_CopySelected_MRAFR_Click(s, event):
  # This function will DUPLICATE the currently selected item (including the questions), AFTER confirmation from user
  #! Note: 17/02/2026: Been re-written to use 'createNewMRA_BasedOnCurrent' function instead as simpler 
  #!       Only diff here is that we want to auto-select this item in the DataGrid after creation. 
  #!       Our new function will return the mraID of the added row (or -1 if error creating header, and -2 if error copying questions) so we can use this to select the new item in the grid after creation.
   
  if dg_MRAFR.SelectedIndex == -1:
    MessageBox.Show("Nothing selected to copy!", "Error: Duplicate Selected item...")
    return
  
  initialConfirmation = "This should only be used for correcting a submitted (completed) MRA.\n\nAre you sure you want to continue?"
  myResult = MessageBox.Show(initialConfirmation, "Duplicate Matter Risk Assessment - confirm...", MessageBoxButtons.YesNo)
  
  if myResult == DialogResult.No:
    return
  
  idItemToCopy = dg_MRAFR.SelectedItem['mraID']
  # if Type is FileReview, then use function to copy File Review instead of MRA (NB: we may want to re-visit this in future as currently this will allow copying of File Reviews, but we may want to restrict this just to MRA?
  if 'File Review' in dg_MRAFR.SelectedItem['Type']:
    newID = createNewFR_BasedOnCurrent(idItemToCopy=idItemToCopy, entRef=_tikitEntity, matNo=_tikitMatter)  
  
    if int(newID) > 0:
      dg_MRAFR_Refresh_and_Reselect(mraIDToReselect=newID)

  else:
    newMRAid = createNewMRA_BasedOnCurrent(idItemToCopy=idItemToCopy, entRef=_tikitEntity, matNo=_tikitMatter)

    if int(newMRAid) > 0:
      dg_MRAFR_Refresh_and_Reselect(mraIDToReselect=newMRAid)

  return  

  
def dg_MRAFR_ViewSelected(s, event):
  # This function will go to the corresponding 'Matter Risk Assessment' or 'File Review' tab, and load questsions and current answers in Read-Only mode
  #! Linked to XAML control.event: btn_View_MRAFR.Click

  if dg_MRAFR.SelectedIndex == -1:
    MessageBox.Show("Nothing selected to Edit!", "Error: Edit selected item...")
    return  

  tmpType = dg_MRAFR.SelectedItem['Type']
  
  if 'File Review' in tmpType:
    # is a FR...
    LOAD_EDIT_FR_PAGE(readOnly=True)
  else:
    # is a MRA... first need to load up the 'Questions' tab and then select the tab
    LOAD_EDIT_MRA_PAGE(readOnly=True)
    
  ti_Main.Visibility = Visibility.Collapsed
  return
  
  
def btn_Edit_MRAFR_Click(s, event):
  # This function will load the 'Questions' tab for the selected item 

  # Reset session token
  global token1 
  token1 = 0

  if dg_MRAFR.SelectedIndex == -1:
    MessageBox.Show("Nothing selected to Edit!", "Error: Edit selected item...")
    return  
  if dg_MRAFR.SelectedItem['Status'] == 'Complete':
    MessageBox.Show("You cannot edit an item already marked as 'Complete'!", "Error: Edit selected item...")
    return
  
  tmpType = dg_MRAFR.SelectedItem['Type']
  #MessageBox.Show("tmpType: " + str(tmpType) + "\ntmpName: " + str(tmpName) + "\ntmpID: " + str(tmpID), "DEBUG: Test Selected Values")
  
  if 'Matter Risk Assessment' in tmpType:
    if _tikitUser != tb_FERef.Text:
      if canUserApproveFeeEarner(UserToCheck = _tikitUser, FeeEarner = tb_FERef.Text) == False and canUserApproveFeeEarner(UserToCheck = tb_FERef.Text, FeeEarner = _tikitUser) == False:
        MessageBox.Show("Only the matter Fee Earner, or the Fee Earners' Approver(s) can edit!", "Error: Edit selected item...")
        return

  if 'File Review' in tmpType:
    LOAD_EDIT_FR_PAGE(readOnly=False)
  else:
    LOAD_EDIT_MRA_PAGE(readOnly=False)
    
  ti_Main.Visibility = Visibility.Collapsed
  return


def LOAD_EDIT_MRA_PAGE(readOnly=False):
  # is a MRA... first need to load up the 'Questions' tab and then select the tab
  lbl_MRA_Name.Content = dg_MRAFR.SelectedItem['Name']
  lbl_MRA_ID.Content = dg_MRAFR.SelectedItem['mraID']
  lbl_ScoreTrigger_High.Content = dg_MRAFR.SelectedItem['ScoreTriggerHigh']
  lbl_ScoreTrigger_Medium.Content = dg_MRAFR.SelectedItem['ScoreTriggerMedium']
  lbl_MRA_TemplateID.Content = dg_MRAFR.SelectedItem['TemplateID']
  lbl_MRA_Status.Content = dg_MRAFR.SelectedItem['Status']

  # get answerlist in memory for this template
  MRA_load_Answers_toMemory()
  # load questions and answers for this MRA into datagrid
  MRA_load_Questions_DataGrid()

  # do we need to update score as well (did on old version but may want to check)
  #MRA_RecalcTotalScore()
  
  if readOnly == True:
    # if we're just viewing, then disable all answer options and hide 'Save' buttons
    btn_MRA_BackToOverview.Visibility = Visibility.Visible
    btn_MRA_Submit.Visibility = Visibility.Collapsed
    btn_MRA_SaveAsDraft.Visibility = Visibility.Collapsed
    #btn_MRA_SaveAnswer.IsEnabled = False
    lbl_TimeEntered.Content = ''
    grp_MRA_SelectedQuestionArea.IsEnabled = False
  else:
    # in 'Edit' mode, so show / hide 'Save' buttons accordingly
    btn_MRA_BackToOverview.Visibility = Visibility.Collapsed
    btn_MRA_Submit.Visibility = Visibility.Visible
    btn_MRA_SaveAsDraft.Visibility = Visibility.Visible
    #btn_MRA_SaveAnswer.IsEnabled = True
    chk_MRA_AutoSelectNext.IsEnabled = True
    lbl_TimeEntered.Content = runSQL("SELECT CONVERT(NVARCHAR, GETDATE(), 121)")
    grp_MRA_SelectedQuestionArea.IsEnabled = True

  # also - if current user is a Risk user (eg: has risk key), display additional columns (score / email comment)
  if _tikitUser in RiskAndITUsers:
    dg_MRA.Columns[7].Visibility = Visibility.Visible   # Email comment column
    dg_MRA.Columns[5].Visibility = Visibility.Visible   # Score column
    stk_RiskInfo.Visibility = Visibility.Visible
  else:
    dg_MRA.Columns[7].Visibility = Visibility.Collapsed
    dg_MRA.Columns[5].Visibility = Visibility.Collapsed
    stk_RiskInfo.Visibility = Visibility.Collapsed

  populate_MRA_DaysUntilLocked(expiryDate=dg_MRAFR.SelectedItem['Expiry'])
  ti_MRA.Visibility = Visibility.Visible
  ti_MRA.IsSelected = True


def LOAD_EDIT_FR_PAGE(readOnly=False):

  global UserCanReviewFiles
  if UserCanReviewFiles == False:
    MessageBox.Show("Only the Fee Earner's HOD, or the Fee Earners' Approver(s) can edit!", "Error: Edit selected item...")
    return
  
  # is a FR...
  lbl_FR_Name.Content = dg_MRAFR.SelectedItem['Name']
  lbl_FR_ID.Content = dg_MRAFR.SelectedItem['mraID']
  
  # refresh File Review datagrid
  refresh_FR(None, None)

  if readOnly == True:
    # set 'view' mode option
    opt_EditModeFR.IsChecked = False
    # enable Submit button
    btn_FR_Submit.IsEnabled = False 
    # enable Answer Option buttons
    opt_Yes.IsEnabled = False
    opt_No.IsEnabled = False
    opt_NA.IsEnabled = False
    lbl_TimeEnteredFR.Content = ''
  else:
    # set 'edit' mode option    
    opt_EditModeFR.IsChecked = True
    # enable Submit button
    btn_FR_Submit.IsEnabled = True 
    # enable Answer Option buttons
    opt_Yes.IsEnabled = True
    opt_No.IsEnabled = True
    opt_NA.IsEnabled = True
    lbl_TimeEnteredFR.Content = runSQL("SELECT CONVERT(NVARCHAR, GETDATE(), 121)")

  #dg_FR.IsEnabled = True
  # show the 'auto go to next Question' and go to FR tab
  chk_FR_AutoSelectNext.Visibility = Visibility.Visible
  ti_FR.Visibility = Visibility.Visible
  ti_FR.IsSelected = True


def btn_HOD_Approval_MRA_Click(s, event):
  #! This is the 'Approval' button on the main 'Overview' tab (eg: so this function is tied to the DataGrid 'dg_MRAFR' and not the 'Questions' tab).
  # - The button on the 'Questions' tab is linked to the 'btn_HOD_Approval_MRA1_Click' function below.
  # New button added for HOD to approve a High Risk MRA (no checks are made here as to whether user is a HOD because this is handled onload (eg: if user is not HOD, button remains disabled)
  if dg_MRAFR.SelectedIndex == -1:
    MessageBox.Show("Nothing selected to 'Approve'!", "Error: HOD Approval for High Risk matter...")
    return   

  tmpIndex = dg_MRAFR.SelectedIndex
  returnVal = HOD_Approves_Item(myOV_ID = dg_MRAFR.SelectedItem['mraID'], 
                               myEntRef = _tikitEntity, myMatNo = _tikitMatter, 
                              myMRADesc = dg_MRAFR.SelectedItem['Name'])

  if returnVal == 1:
    dg_MRAFR_Refresh(s, event)
    # and select exiting item
    dg_MRAFR.SelectedIndex = tmpIndex
    # and scroll into view
    dg_MRAFR.ScrollIntoView(dg_MRAFR.SelectedItem)
  return


def btn_HOD_Approval_MRA1_Click(s, event):
  # New button added for HOD to approve a High Risk MRA (no checks are made here as to whether user is a HOD because this is handled onload (eg: if user is not HOD, button remains disabled)
  # Note: This button is the one on the actual 'Edit' page (rather than previous function that's for the button on the 'Overview' tab)

  returnVal = HOD_Approves_Item(myOV_ID = lbl_MRA_ID.Content, 
                               myEntRef = _tikitEntity, myMatNo = _tikitMatter, 
                              myMRADesc = lbl_MRA_Name.Content)

  if returnVal == 1:
    btn_HOD_Approval_MRA1.IsEnabled = False
    ti_Main.Visibility = Visibility.Visible
    ti_Main.IsSelected = True
    ti_MRA.Visibility = Visibility.Collapsed
    # refresh main overview datagrid (as we've updated 'Approved' status)
    dg_MRAFR_Refresh(s, event)
  return



def getSQLDate(varDate):
  #Converts the passed varDate into SQL version date (YYYY-MM-DD)

  newDate = ''
  tmpDate = ''
  tmpDay = ''
  tmpMonth = ''
  tmpYear = ''
  mySplit = []
  finalStr = ''
  canContinue = False

  # If passed value is of 'DateTime' then convert to string
  if isinstance(varDate, DateTime) == True:
    tmpDate = varDate.ToString()
    canContinue = True

  # else if already a string, assingn passed date directly into newDate 
  elif isinstance(varDate, str) == True:
    tmpDate = varDate                       #datetime.datetime(varDate) '1/1/2020'
    canContinue = True

  if canContinue == True:
    # now to strip out the time element
    mySplit = []
    mySplit = tmpDate.split(' ')
    newDate = mySplit[0]

    #MessageBox.Show('newDate is ' + newDate)
    mySplit = []

    if len(newDate) >= 8:
      mySplit = newDate.split('/')

      tmpDay = mySplit[0]             #newDate.strftime("%d")
      tmpMonth = mySplit[1]           #newDate.strftime("%m")
      tmpYear = mySplit[2]            #newDate.strftime("%Y")

      testStr = str(tmpYear) + '-' + str(tmpMonth) + '-' + str(tmpDay)
        #MessageBox.Show('Original: ' + str(varDate) + '\nFinal: ' + testStr)
        #newDate1 = datetime.datetime(int(tmpYear), int(tmpMonth), int(tmpDay))
        #finalStr = newDate1.strftime("%Y-%m-%d")
      finalStr = testStr

    return finalStr

def dg_MRAFR_DeleteSelected(s, event):
  # This function will delete the selected Matter Risk Assessment/File Review, and any questions associated to it
  # # # #  NB: If item is a 'File Review', then need to check for any 'Corrective Actions'... and what to do if any exists (do we delete them too, or leave them?
  # # # #      Currently we are KEEPING any Corrective Actions, but we ought to ask Johan if one expects them to be deleted too.
  # # # #      Also, it may be desirable that the 'Delete' button is disabled for all users except IT/Risk... again, ask Johan
  
  if dg_MRAFR.SelectedIndex == -1:
    MessageBox.Show("Nothing selected to delete!", "Error: Delete Selected Matter Risk Assessment...")
    return

  # let's first ask for confirmation before blindly deleting the item (and associated questions)
  result = MessageBox.Show("Are you sure you want to delete the selected item and its associated questions?", "Confirm Delete", MessageBoxButtons.YesNo)
  if result != DialogResult.Yes:
    return

  # First get the ID, as we'll also want to delete questions using this ID
  tmpID = dg_MRAFR.SelectedItem['mraID'] 
  if 'File Review' in dg_MRAFR.SelectedItem['Type']:
    tmpType = 'File Review'
    deleteSQL = "DELETE FROM Usr_MRA_Overview WHERE ID = {id} AND EntityRef = '{entRef}' AND MatterNo = {matNo}".format(id=tmpID, entRef=_tikitEntity, matNo=_tikitMatter)
    questionsTable = 'Usr_MRA_Detail'
    questionsIDfield = 'OV_ID'
  else:
    tmpType = 'Matter Risk Assessment'
    deleteSQL = "DELETE FROM Usr_MRAv2_MatterHeader WHERE mraID = {id} AND EntityRef = '{entRef}' AND MatterNo = {matNo}".format(id=tmpID, entRef=_tikitEntity, matNo=_tikitMatter)
    questionsTable = 'Usr_MRAv2_MatterDetails'
    questionsIDfield = 'mraID'

  try:
    _tikitResolver.Resolve("[SQL: {0}]".format(deleteSQL))
  except Exception as e:
    MessageBox.Show("There was an error deleting the selected item, using SQL:\n{0}\n\nError details:\n{1}".format(deleteSQL, str(e)), "Error: Delete Selected {2}...".format(tmpType))
    return
  
  # if we get here, we have deleted the 'header' row, so now need to delete any Questions associated to this item 
  deleteQ_sql = """DELETE FROM {questionsTable} 
                    WHERE {questionsIDfield} = {id} 
                    AND EntityRef = '{entRef}' AND MatterNo = {matNo}""".format(questionsTable=questionsTable, 
                                                                                questionsIDfield=questionsIDfield, 
                                                                                id=tmpID, 
                                                                                entRef=_tikitEntity, 
                                                                                matNo=_tikitMatter)

  try:
    _tikitResolver.Resolve("[SQL: {0}]".format(deleteQ_sql))
    MessageBox.Show("Selected {0} and associated Questions have been deleted successfully.".format(tmpType), "Delete successful")
  except Exception as e:
    MessageBox.Show("There was an error deleting the associated Questions for the selected item, using SQL:\n{0}\n\nError details:\n{1}".format(deleteQ_sql, str(e)), "Error: Delete Questions for Selected {2}...".format(tmpType))
    return
  return


# # C O R R E C T I V E   A C T I O N S   (O V E R V I E W)  # #
class CA(object):                                                                     
  def __init__(self, myID, myQText, myReviewerName, myCANeeded, myCATaken, myComplete, myDueBy, myFRID, myReviewerID, myQNum, myFRName):
    self.ID = myID
    self.CA_FRQText = myQText
    self.CA_Needed = myCANeeded
    self.CA_Taken = myCATaken
    self.CA_Complete = myComplete
    self.CA_DueBy = myDueBy
    self.CA_Reviewer = myReviewerName
    self.FR_ID = myFRID
    self.FR_U_ID = myReviewerID
    self.FR_QNum = myQNum
    self.FR_Name = myFRName
    #self.CA_CompleteYN = 'No' if myComplete == 0 else 'Yes'
    return
    
  def __getitem__(self, index):
    if index == 'CA_ID':
      return self.ID
    elif index == 'FR Name and Q':
      return self.CA_FRQText
    elif index == 'CA Needed':
      return self.CA_Needed
    elif index == 'CA Taken':
      return self.CA_Taken
    elif index == 'Completed':
      if self.CA_Complete == False:
        return 0
      elif self.CA_Complete == True:
        return 1
      else:
        return self.CA_Complete
    elif index == 'DueBy':
      return self.CA_DueBy
    elif index == 'Reviewer':
      return self.CA_Reviewer
    elif index == 'FR ID':
      return self.FR_ID
    elif index == 'Reviewer ID':
      return self.FR_U_ID
    elif index == 'FR Q Num':
      return self.FR_QNum
    elif index == 'FR Name':
      return self.FR_Name
    else:
      return 'Unrecognised index provided ({0})'.format(index)
    return
    
def dgCA_Overview_Refresh(s, event):
  # This function populates the Corrective Actions DataGrid (on main page)

  myEntity = _tikitEntity
  myMatNo = _tikitMatter

  mySQL = """SELECT '0-CA ID' = MA.ID, '1-QText' = CONCAT('Q', TQ.DisplayOrder, ': ', TQ.QuestionText), 
                  '2-Reviewer' = '(' + U.Code + ') ' + U.FullName, '3-CA Needed' = MA.CorrActionNeeded, 
                  '4-CA Taken' = MA.CorrActionTaken, '5-Complete' = MA.AuditPass, '6-Due By' = MA.NextAuditDate, 
                  '7-FR ID' = MRAO.ID, '8-FR Reviewer' = ISNULL(MRAO.FR_Reviewer, ''), '9-Q Num ID' = FRD.ID,
                  '10-LocalName' = MRAO.LocalName 
            FROM Usr_MRA_Detail FRD 
                LEFT OUTER JOIN Usr_MRA_Overview MRAO ON FRD.OV_ID = MRAO.ID 
                LEFT OUTER JOIN Matter_Audit MA ON FRD.CorrActionID = MA.ID 
                LEFT OUTER JOIN Usr_MRA_TemplateTypes TT ON MRAO.TypeID = TT.TypeID 
                LEFT OUTER JOIN Users U ON MRAO.FR_Reviewer = U.Code 
                LEFT OUTER JOIN Usr_MRA_TemplateQs TQ ON FRD.QuestionID = TQ.ID 
            WHERE MA.EntityRef = '{0}' AND MA.MatterNo = {1} AND TT.Is_MRA = 'N' """.format(myEntity, myMatNo)

  if opt_CA_ViewNotComplete.IsChecked == True:
    mySQL += "AND MA.AuditPass = 0 "
  elif opt_CA_ViewComplete.IsChecked == True:
    mySQL += "AND MA.AuditPass = 1 "

  tmpItem = []
  _tikitDbAccess.Open(mySQL)
  
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          iID = 0 if dr.IsDBNull(0) else dr.GetValue(0)
          iQText = '' if dr.IsDBNull(1) else dr.GetString(1)
          iReviewerName = '' if dr.IsDBNull(2) else dr.GetString(2)
          iCANeeded = '' if dr.IsDBNull(3) else dr.GetString(3)
          iCATaken = '' if dr.IsDBNull(4) else dr.GetString(4)
          iComplete = 0 if dr.IsDBNull(5) else dr.GetValue(5)
          iDueBy = 0 if dr.IsDBNull(6) else dr.GetValue(6)
          iFRID = 0 if dr.IsDBNull(7) else dr.GetValue(7)
          iReviewerID = 0 if dr.IsDBNull(8) else dr.GetValue(8)
          iQNum = 0 if dr.IsDBNull(9) else dr.GetValue(9)
          iFRName = '' if dr.IsDBNull(10) else dr.GetString(10)

          tmpItem.append(CA(iID, iQText, iReviewerName, iCANeeded, iCATaken, iComplete, iDueBy, iFRID, iReviewerID, iQNum, iFRName))

    dr.Close()
  
  #close db connection
  _tikitDbAccess.Close()
  dgCA_Overview.ItemsSource = tmpItem

  if dgCA_Overview.Items.Count > 0:
    dgCA_Overview.Visibility = Visibility.Visible
    tb_NoCAs.Visibility = Visibility.Hidden
    #btn_Mark_CA_Complete.IsEnabled = True
    #btn_View_CA_onFR.IsEnabled = True
  else:
    tb_NoCAs.Visibility = Visibility.Visible
    dgCA_Overview.Visibility = Visibility.Hidden
  
  btn_Mark_CA_Complete.IsEnabled = False
  btn_View_CA_onFR.IsEnabled = False

  # call function to set enabled state of 'Notify Reviewer' button (if all outstanding CAs have an 'ActionTaken' note added to each)
  btn_UpdateReviewerWithActionTaken_SetEnabled()

  return

def btn_UpdateReviewerWithActionTaken_SetEnabled():
  # This function will set the enabled state of the 'Notify Reviewer' button, based on whether there are any incomplete CAs with a note in the 'Action Taken' field
  #! Called by: dgCA_Overview_Refresh, dgCA_Overview_CellEditEnding

  # form SQL to get the count of incomplete Corrective Actions for matter, and run (this is: CA's not yet ticked off/marked as complete)
  count_CA_withFEnote_sql = """SELECT COUNT(QuestionID) FROM Usr_MRA_Detail MRAD 
                                LEFT OUTER JOIN Matter_Audit MA ON MRAD.CorrActionID = MA.ID 
                                WHERE MA.AuditPass = 0 
                                AND MA.EntityRef = '{0}' AND MA.MatterNo = {1}
                                AND LEN(ISNULL(MA.CorrActionTaken, '')) > 1 """.format(_tikitEntity, _tikitMatter)
  count_CA_withFEnote = runSQL(count_CA_withFEnote_sql)

  if int(count_CA_withFEnote) > 0:
    btn_UpdateReviewerWithActionTaken.IsEnabled = True
    btn_UpdateReviewerWithActionTaken.Background = Brushes.LightGreen
  else:
    btn_UpdateReviewerWithActionTaken.IsEnabled = False
    btn_UpdateReviewerWithActionTaken.Background = Brushes.AliceBlue
  # I seem to recall talk of always enabling this button, but just have a message box pop-up if it doesn't make sense in the context of the current matter
  # but I'd prefer what we're doing here, as it gives a visual indication of whether the button is active or not (and why)

  # OLD LOGIC:
  # form SQL to get the count of incomplete Corrective Actions for matter, and run (this is: CA's not yet ticked off/marked as complete)
  #CA_NotComplete_SQL = """SELECT COUNT(QuestionID) FROM Usr_MRA_Detail MRAD 
  #                        LEFT OUTER JOIN Matter_Audit MA ON MRAD.CorrActionID = MA.ID 
  #                        WHERE MA.AuditPass = 0 
  #                        AND MA.EntityRef = '{0}' AND MA.MatterNo = {1}""".format(myEntity, myMatNo)
  #count_CA_NotComplete = runSQL(CA_NotComplete_SQL)
  #MessageBox.Show("Entering function\n(ovID={0}, callingFrom={1})\n\nCountIfIncompleteCAs: {2}".format(ovID, callingFrom, countOfIncompleteCAs), "DEBUGGING - FR_checkForOSca_andFinalise")

  #  highlight the Notify Review Actions Complete button when all Actions Taken are completed
  # formally, this was stating "And CorrActionTaken = ''" But now added a min length of 2
  #count_CA_withFEnote_sql = """SELECT COUNT(QuestionID) FROM Usr_MRA_Detail MRAD 
  #                              LEFT OUTER JOIN Matter_Audit MA ON MRAD.CorrActionID = MA.ID 
  #                              WHERE MA.AuditPass = 0 
  #                              AND MA.EntityRef = '{0}' AND MA.MatterNo = {1}
  #                              AND LEN(ISNULL(MA.CorrActionTaken, '')) > 1 """.format(myEntity, myMatNo)
  # ^ Note: this is 'Incomplete CAs' (same as above) but with additional 'WHERE' corrective action has been taken (ie: not blank)
  #count_CA_withFEnote = runSQL(count_CA_withFEnote_sql)
  
  #count_FRwithFE_SQL = """SELECT COUNT(MRAO.ID) 
  #                        FROM Usr_MRA_Overview MRAO
  #                          JOIN Usr_MRA_TemplateTypes TT ON MRAO.TypeID = TT.TypeID
  #                        WHERE MRAO.EntityRef = '{0}' AND MatterNo = {1}
  #                        AND MRAO.Status = 'With FE' AND TT.Is_MRA = 'N' """.format(myEntity, myMatNo)
  #                        
  #count_FRwithFE = runSQL(count_FRwithFE_SQL)

  # if there's 1 or more incomplete CA's, show the button to 'Notify File Reviewer action taken' (else, disable it)
  #if int(count_CA_NotComplete) > 0 and int(count_FRwithFE) > 0 and int(count_CA_withFEnote) > 0:
  # ^ old logic was checking for count of incomplete CAs greater than 0, and had to have status of 'With FE', and also had to have a note in the 'Action Taken' field
  #   But seems overly complex to me... also our count 'CA with FE note' is counting INCOMPLETE CAs which have a text entered in 'Action Taken' field
  #     - which means we don't need the first SQL bit (count_CA_NotComplete)
  #     - and begs the question: do we need to check for 'With FE' status?
  #       - I think not, as it's only 'complete' once all CA's have been marked as such (AuditPass = 1)
  #  btn_UpdateReviewerWithActionTaken.IsEnabled = True
  #  btn_UpdateReviewerWithActionTaken.Background = Brushes.LightGreen
  #else:
  #  btn_UpdateReviewerWithActionTaken.IsEnabled = False
  #  btn_UpdateReviewerWithActionTaken.Background = Brushes.AliceBlue

  return


def dgCA_Overview_SelectionChanged(s, event):
  # Linked to XAML control.event: dgCA_Overview.SelectionChanged
  # This function will populate the 'Current_CA' text fields (with data for current selected row), and set the 
  # enabled state of the 'Toggle Complete' and 'View on File Review' buttons

  # if nothing selected
  if dgCA_Overview.SelectedIndex == -1:
    # set text boxes to null / empty values and disable action buttons (toggle complete/view on FR)
    tb_Current_CA.Text = 'x'
    tb_Current_CA_Complete.Text = ''
    btn_Mark_CA_Complete.IsEnabled = False
    btn_View_CA_onFR.IsEnabled = False
    tb_CATaken.Text = ''
  else:
    #MessageBox.Show("Value of 'CA Taken': {0}".format(dgCA_Overview.SelectedItem['CA Taken']), "DEBUGGING")
    # set text boxes to current (selected) row values and enable action buttons (toggle complete/view on FR)
    tb_Current_CA.Text = str(dgCA_Overview.SelectedItem['CA_ID'])
    tb_Current_CA_Complete.Text = str(dgCA_Overview.SelectedItem['Completed'])
    tb_CATaken.Text = str(dgCA_Overview.SelectedItem['CA Taken'])
    btn_Mark_CA_Complete.IsEnabled = True
    btn_View_CA_onFR.IsEnabled = True  
  return


def dgCA_Overview_CellEditEnding(s, event):
  # Linked to XAML control.event: dgCA_Overview.CellEditEnding 
  # Only the 'Complete' checkbox can be changed in this DataGrid, and this function commits the change to the database

  # if nothing selected exit function now
  if dgCA_Overview.SelectedIndex == -1:
    return

  # ADD CODE HERE TO ALLOW EDITING OF 'CORRECTIVE ACTION TAKEN'
  tmpColHeader = event.Column.Header
  caID = dgCA_Overview.SelectedItem['CA_ID']
  caTaken = dgCA_Overview.SelectedItem['CA Taken']
  caComplete = dgCA_Overview.SelectedItem['Completed']

  # if column edited is 'Corrective Action Taken'
  if tmpColHeader == 'Action Taken':
    # may want/need to add other checks in here so that only matter FE or Reviewer can edit??
    # if CA is already marked as 'Complete', don't allow update
    if caComplete == 1:
      MessageBox.Show("Please note that this Corrective Action has been marked as 'Complete' and therefore cannot be edited", "Edit 'Action Taken' not allowed...")
      return

    # if new 'Taken' text is different from old 'Taken' text
    if caTaken != tb_CATaken.Text:
      # form the SQL to update 'Taken' text
      updateSQL = "UPDATE Matter_Audit SET CorrActionTaken = '{0}' WHERE ID = {1}".format(caTaken.replace("'", "''"), caID)
      runSQL(updateSQL)
      #MessageBox.Show("Value is different so updated using SQL: {0}".format(updateSQL), "DEBUGGING")

  # if columnt edited is 'CA Complete' (checkbox)
  if tmpColHeader == 'Complete':
    # if new value is different from old value
    if int(caComplete) != int(tb_Current_CA_Complete.Text):
      # just call the function that already does this
      dgCA_Overview_ToggleComplete(s, event)
      
  # call function to set enabled state of 'Notify Reviewer' button (if all outstanding CAs have an 'ActionTaken' note added to each)
  btn_UpdateReviewerWithActionTaken_SetEnabled()
  return


def dgCA_Overview_ToggleComplete(s, event):
  # This function will mark the selected 'Corrective Action' as 'Complete' - note: only the file reviewer can mark as 'complete'
  # May want to double-check with Johan if that logic is correct - if FE can mark as complete, remove 'if...'

  # if text box holding current 'complete' status for selected row is zero (incomplete)
  if tb_Current_CA_Complete.Text == '0':
    # set 'new' value to complete (1)
    newCompletedVal = 1
  else:
    # item is already complete, so mark as 'incomplete' (0)
    newCompletedVal = 0
  
  # get who did the File Review, and get ID of CA
  tmpItemReviewer = dgCA_Overview.SelectedItem['Reviewer ID']
  currDGId = tb_Current_CA.Text

  # here we are assuming only the File Reviewer can toggle the 'Complete' status  
  if tb_CurrUser.Text != tmpItemReviewer:
    # current user doesn't appear to be the File Reviewer, so alert user and exit
    MessageBox.Show("Only the File Reviewer can toggle the 'Complete' status!", "Error: Toggle 'Complete' status of Corrective Action...")
    return
  else: 
    # generate SQL to toggle 'Complete' (AuditPass)
    tmpSQL = "[SQL: UPDATE Matter_Audit SET AuditPass = {0} WHERE ID = {1}]".format(newCompletedVal, currDGId)
    #MessageBox.Show("tmpSQL: {0}".format(tmpSQL)) 
    try:
      # run the SQL, and refresh the CA overview list
      _tikitResolver.Resolve(tmpSQL)
    except:
      # there was an error running the SQL to toggle 'Complete' status, alert user
      MessageBox.Show("There was an error toggling the 'Complete' status, using SQL:\n{0}".format(tmpSQL), "Error: Toggle 'Complete' status of Corrective Action...")

    # new: check if any outstanding CA and if not, mark FR as complete
    FR_checkForOSca_andFinalise(sender=s, e=event, ovID=int(dgCA_Overview.SelectedItem['FR ID']))

    # finally refresh the Corrective Actions datagrid and if 'ViewAll' selected, find item and re-select it
    dgCA_Overview_Refresh(s, event)
    # ideally want to select same item again after refresh, but can only do so if 'View All' is selected
    # as on the other views, when toggling 'complete' status, it will dissappear from current view
    if opt_CA_ViewAll.IsChecked == True:
      tCount = -1
      for myRow in dgCA_Overview.Items:
        tCount += 1
        if myRow.ID == int(currDGId):        
          dgCA_Overview.SelectedIndex = tCount
          break
    else:
      # otherwise (not 'viewing all'), check to see if there are items, and select first one
      if dgCA_Overview.Items.Count > 0:
        dgCA_Overview.SelectedIndex = 0
  return

  
def dg_CA_Overview_ViewOnFileReview(s, event):
  # This function will go to the 'File Review' tab, and will select the Question to which the CA was added to, and populate bottom area of tab
  
  tmpFRID = dgCA_Overview.SelectedItem['FR ID']
  tmpFRName = _tikitResolver.Resolve("[SQL: SELECT LocalName FROM Usr_MRA_Overview WHERE ID = {0}]".format(tmpFRID))
  tmpQID = dgCA_Overview.SelectedItem['FR Q Num']
  
  lbl_FR_Name.Content = tmpFRName
  lbl_FR_ID.Content = str(tmpFRID)
    
  # refresh File Review datagrid
  refresh_FR(s, event)

  # now select the appropriate Q
  tCount = -1
  for x in dg_FR.Items:
    tCount +=1
    # MessageBox.Show("Desired Q ID to find: " + str(tmpQID) + "\nQ Num on this row: " + str(x.pvID) + "\ntCount: " + str(tCount), "DEBUGGING: Select approp Q...")
    if int(x.pvID) == int(tmpQID):
      dg_FR.SelectedIndex = tCount
      break

  dg_FR.IsEnabled = True
  ti_FR.Visibility = Visibility.Visible
  ti_FR.IsSelected = True  
  return


def FR_checkForOSca_andFinalise(sender, e, ovID = 0, callingFrom = ''):
  # New function added 7th November 2024
  # This will see if there are any outstanding Corrective Actions, and if there are none, will initiate the 'all complete' Task Centre task
  # NB: updated 2/12/2024 so that it will only set to 'Complete' if all questions have an answer (was a chance before that this could set as 'complete'
  # without all questions having an answer)

  #MessageBox.Show("Entering function\n(ovID={0}, callingFrom={1})".format(ovID, callingFrom), "DEBUGGING - FR_checkForOSca_andFinalise")
  # set initial variables
  myEntity = str(_tikitEntity)
  myMatNo = _tikitMatter
  # call our new function to update stats on XAML
  update_FR_Stats(ov_ID=ovID)

  countOfIncompleteCAs = int(tb_TotalOSCAs_FR.Text)
  countOfQuestions = int(tb_TotalQs_FR.Text)
  countAnswered = int(tb_TotalAnswered_FR.Text)
  countOfQsNoAnswer = int(countOfQuestions) - int(countAnswered)


  ## form SQL to get the count of incomplete Corrective Actions for matter, and run
  #countOfIncompleteCAs_SQL = """SELECT COUNT(QuestionID) FROM Usr_MRA_Detail MRAD 
  #                              LEFT OUTER JOIN Matter_Audit MA ON MRAD.CorrActionID = MA.ID WHERE MA.AuditPass = 0 
  #                              AND MA.EntityRef = '{0}' AND MA.MatterNo = {1} AND MRAD.OV_ID = {2}""".format(myEntity, myMatNo, ovID)
  #countOfIncompleteCAs = runSQL(countOfIncompleteCAs_SQL)
  ##MessageBox.Show("Entering function\n(ovID={0}, callingFrom={1})\n\nCountIfIncompleteCAs: {2}".format(ovID, callingFrom, countOfIncompleteCAs), "DEBUGGING - FR_checkForOSca_andFinalise")
  #
  ## create SQL to 'count of questions' and 'count of questions with no answer'
  #countOfQuestions_s = """SELECT COUNT(QuestionID) FROM Usr_MRA_Detail MRAD
  #                        LEFT OUTER JOIN Usr_MRA_Overview MRAO ON MRAD.EntityRef = MRAO.EntityRef AND MRAD.MatterNo = MRAO.MatterNo AND MRAD.OV_ID = MRAO.ID 
  #                        LEFT OUTER JOIN Usr_MRA_TemplateTypes TT ON MRAO.TypeID = TT.TypeID
  #                        WHERE MRAD.EntityRef = '{0}' AND MRAD.MatterNo = {1}  AND MRAD.OV_ID = {2} AND TT.Is_MRA = 'N'""".format(myEntity, myMatNo, ovID)
  #countOfQsNoAnswer_s = """SELECT COUNT(QuestionID) FROM Usr_MRA_Detail MRAD
  #                         LEFT OUTER JOIN Usr_MRA_Overview MRAO ON MRAD.EntityRef = MRAO.EntityRef AND MRAD.MatterNo = MRAO.MatterNo AND MRAD.OV_ID = MRAO.ID 
  #                         LEFT OUTER JOIN Usr_MRA_TemplateTypes TT ON MRAO.TypeID = TT.TypeID
  #                         WHERE MRAD.EntityRef = '{0}' AND MRAD.MatterNo = {1}  AND MRAD.OV_ID = {2} AND ISNULL(tbAnswerText, '') = '' AND TT.Is_MRA = 'N'""".format(myEntity, myMatNo, ovID)
  #
  ## run the SQL to get answer
  #countOfQuestions = runSQL(countOfQuestions_s)
  #countOfQsNoAnswer = runSQL(countOfQsNoAnswer_s)

  # if there's 1 or more incomplete CA's...
  if int(countOfIncompleteCAs) > 0:

    # if count of questions is greater than one
    if int(countOfQuestions) > 0:
      # if the count of questions without an answer is 0
      if int(countOfQsNoAnswer) > 0:
        tmpStatus = 'Draft'
        tmpTriggerText = 'FR_Draft' 
        doSendTaskCentreEmail = False
      else:
        # all questions have an answer, but there are CA's outstanding (not 'completed')
        # set Status to state 'With FE'...
        tmpStatus = 'With FE'
        tmpTriggerText = 'FR_CorrectiveActions_WithFE'
        # only set 'send TC email' to True if calling function was the initial 'Submit' button click (when editing the FR)
        doSendTaskCentreEmail = True if callingFrom == 'btn_FR_Submit_Click' else False
        # ^ why? if we're meant to be auto-completing, then why would we NOT want to send the email?
    else:
      # there don't appear to be any questions for this File Review, so mark as 
      tmpStatus = 'Err-NoQ HasCA'
      tmpTriggerText = '' 
      doSendTaskCentreEmail = False
  else:
  
    # if there's more than 0 questions
    if int(countOfQuestions) > 0:
      # if count of quesions without an answer is zero - set status to 'Complete' and send TC task email
      if int(countOfQsNoAnswer) == 0:
        tmpStatus = 'Complete'
        tmpTriggerText = 'FR_Complete' 
        doSendTaskCentreEmail = True  
      else:
        # there's at least one Question without an answer, so set to 'Draft' and do NOT send TC email
        tmpStatus = 'Draft'
        tmpTriggerText = 'FR_Draft' 
        doSendTaskCentreEmail = False
    else:
      # there are NO questions for File Review, set Status to 'Complete' (no Questions and no CAs)...
      tmpStatus = 'Complete'
      tmpTriggerText = 'FR_Complete' 
      doSendTaskCentreEmail = True

  #MessageBox.Show("countIncompleteCAs: {0}\ntmpStatus: {1}\ntmpTriggerText: {2}\nSendTaskCentreEmail: {3}".format(countOfIncompleteCAs, tmpStatus, tmpTriggerText, doSendTaskCentreEmail), "DEBUGGING = FR_checkForOSca_andFinalise(ovID={0}, callingFrom={1})".format(ovID, callingFrom))
  # update main 'Overview' table with new 'Status'
  runSQL("UPDATE Usr_MRA_Overview SET Status = '{0}' WHERE ID = {1}".format(tmpStatus, ovID))

  # to work-around an issue that could arrise by a reviewer ticking off Corrective Actions (eg: this function is called every time a corrective
  # action is toggled, and therefore without intervention, would mean that we email the FE [and reviewer] after every update)
  # We intrtoduced the 'doSendTaskCentreEmail' parameter, to allow us to selectively chose WHEN to send out TC email
  if doSendTaskCentreEmail == True:
    # now set variables to pass into 'mra_events' table
    tmpOurRef = "{0}{1}/{2}".format(myEntity[0:3], myEntity[11:15], myMatNo)
    tmpMatDesc = runSQL(codeToRun="SELECT Description FROM Matters WHERE EntityRef = '{0}' AND Number = {1}".format(myEntity, myMatNo), apostropheHandle=1)
    tmpClName = runSQL(codeToRun="SELECT LegalName FROM Entities WHERE Code = '{0}'".format(myEntity), apostropheHandle=1)
    # email to = Matter Fee Earner | email CC = current user
    tmpEmailTo = runSQL("SELECT EMailExternal FROM Users WHERE Code = (SELECT FeeEarnerRef FROM Matters WHERE EntityRef = '{0}' AND Number = {1})".format(myEntity, myMatNo))
    tmpToUserName = runSQL(codeToRun="SELECT ISNULL(Forename, FullName) FROM Users WHERE Code = (SELECT FeeEarnerRef FROM Matters WHERE EntityRef = '{0}' AND Number = {1})".format(myEntity, myMatNo), apostropheHandle=1)
    tmpEmailCC = runSQL("SELECT EMailExternal FROM Users WHERE Code = '{0}'".format(_tikitUser))
    tmpLocalName =  runSQL("SELECT ISNULL(LocalName, 'File Review') FROM Usr_MRA_Overview WHERE ID = {0}".format(ovID))

    # Insert a record into MRA Events table to trigger email to FE
    insert_into_MRAEvents(userRef=_tikitUser, triggerText=tmpTriggerText, ov_ID=ovID, 
                          emailTo=tmpEmailTo, emailCC=tmpEmailCC, toUserName=tmpToUserName, 
                          ourRef=tmpOurRef, matterDesc=tmpMatDesc, clientName=tmpClName, 
                          addtl1=tmpLocalName, addtl2=countOfIncompleteCAs)

  # finally, refresh 'overview' list
  dg_MRAFR_Refresh(sender, e)

  return
# # # #   END OF:   O V E R V I E W    TAB   # # # # 


# # # #    M A T T E R   R I S K   A S S E S S M E N T    TAB  # # # #
# # V2: Updated to use new model (MVVM approach) - Data is loaded into model in UI, and only when user clicks 'Save as Draft' or 'Submit' do we then
#       write back to the database - this is a much cleaner approach, and also means we don't have to run multiple updates to the database as user is making changes (eg: selecting answers, adding comments etc) 
#       One difference that the Practice version doesn't have (on the 'Preview MRA' tab), is the 'Comments' field (as that's just a 'practice' screen for testing changes without commiting and checking on live matter)

class NotifyBase(INotifyPropertyChanged):
  # This creates add_PropertyChanged/remove_PropertyChanged automatically
  def __init__(self):
    # store delegates that WPF adds via add_PropertyChanged
    self._pc_handlers = []

  # .NET event accessor: WPF calls this when binding subscribes
  def add_PropertyChanged(self, handler):
    if handler is None:
      return
    self._pc_handlers.append(handler)

  # .NET event accessor: WPF calls this when binding unsubscribes
  def remove_PropertyChanged(self, handler):
    if handler is None:
      return
    # remove first matching instance
    for i in range(len(self._pc_handlers) - 1, -1, -1):
      if self._pc_handlers[i] == handler:
        del self._pc_handlers[i]
        break

  
  def _raise(self, prop_name):
    if not self._pc_handlers:
      return
    args = PropertyChangedEventArgs(prop_name)
    # iterate over a copy in case handlers mutate subscriptions
    for h in list(self._pc_handlers):
      h(self, args)

# -------------------------
# Answer model
# -------------------------
class AnswerItem(object):
  def __init__(self, answer_id, text, email_comment="", score=0):
    self.AnswerID = int(answer_id) if answer_id is not None else None
    self.AnswerText = text or ""
    self.EmailComment = email_comment or ""
    self.Score = int(score) if score is not None else 0

  def __repr__(self):
    return self.AnswerText

# -------------------------
# Question model (bindable)
# -------------------------
class MatterQuestionItem(NotifyBase):
  def __init__(self, group_name, order_no, qid, qtext, answers):
    NotifyBase.__init__(self)
    self.QuestionGroup = group_name or ""
    self.QuestionOrder = int(order_no) if order_no is not None else 0
    self.QuestionID = int(qid) if qid is not None else None
    self.QuestionText = qtext or ""
    self.AvailableAnswers = answers or []
    self._SelectedAnswer = None  # <-- bind target
    self._UserComment = ""   # free-form per question

  @property
  def UserComment(self):
    return self._UserComment

  @UserComment.setter
  def UserComment(self, value):
    v = "" if value is None else str(value)
    if self._UserComment == v:
      return
    self._UserComment = v
    self._raise("UserComment")

  @property
  def SelectedAnswer(self):
    return self._SelectedAnswer

  @SelectedAnswer.setter
  def SelectedAnswer(self, value):
    # value should be an AnswerItem (or None)
    if self._SelectedAnswer == value:
      return
    self._SelectedAnswer = value

    # Notify all dependents
    self._raise("SelectedAnswer")
    self._raise("SelectedAnswerID")
    self._raise("SelectedAnswerText")
    self._raise("SelectedAnswerScore")
    self._raise("SelectedAnswerEmailComment")

  # Computed (read-only) convenience properties
  @property
  def SelectedAnswerID(self):
    return None if self._SelectedAnswer is None else self._SelectedAnswer.AnswerID

  @property
  def SelectedAnswerText(self):
    return "" if self._SelectedAnswer is None else self._SelectedAnswer.AnswerText

  @property
  def SelectedAnswerScore(self):
    return 0 if self._SelectedAnswer is None else self._SelectedAnswer.Score

  @property
  def SelectedAnswerEmailComment(self):
    return "" if self._SelectedAnswer is None else self._SelectedAnswer.EmailComment
  

def _get_preview_current_question():
  view = CollectionViewSource.GetDefaultView(dg_MRA.ItemsSource)
  if view is None:
    return None
  return view.CurrentItem

def to_int(x, default=0):
  if is_dbnull(x):
    return default
  try:
    return Convert.ToInt32(x)
  except:
    try:
      return Convert.ToInt32(str(x))
    except:
      return default

def sql_escape(s):
  # Very basic single-quote escaping for building SQL strings.
  if s is None:
    return ""
  return str(s).replace("'", "''")

def is_dbnull(x):
  try:
    return x is None or x == DBNull.Value
  except:
    return x is None
  
def MRA_load_Answers_toMemory():
  global MRA_ANSWERS_BY_QID
  MRA_ANSWERS_BY_QID = {}

  # This differs from the 'Practice' version as we're working with 'live' data that user may have entered already
  # Therefore, need to get TemplateID from 'Usr_MRAv2_MatterHeader' table, and then all answers to Q's get stored in 'Usr_MRAv2_MatterDetails'

  mySQL = """SELECT T.QuestionID, Ans.AnswerID, Ans.AnswerText, Ans.EmailComment, T.Score
             FROM Usr_MRAv2_Templates T
                JOIN Usr_MRAv2_Answer Ans ON T.AnswerID = Ans.AnswerID
             WHERE T.TemplateID = {0}
             ORDER BY T.QuestionID, T.AnswerOrder;""".format(lbl_MRA_TemplateID.Content)

  _tikitDbAccess.Open(mySQL)
  dr = _tikitDbAccess._dr
  if dr is not None and dr.HasRows:
    while dr.Read():
      qid = to_int(dr.GetValue(0))
      aid = to_int(dr.GetValue(1))
      text = "" if dr.IsDBNull(2) else dr.GetString(2)
      ec = "" if dr.IsDBNull(3) else dr.GetString(3)
      score = 0 if dr.IsDBNull(4) else to_int(dr.GetValue(4))

      item = AnswerItem(aid, text, ec, score)
      MRA_ANSWERS_BY_QID.setdefault(qid, []).append(item)

    dr.Close()
  _tikitDbAccess.Close()
  return

def MRA_load_Questions_DataGrid():
  # This function will populate the Matter Risk Assessment Preview datagrid
  #MessageBox.Show("Start - getting group ID", "Refreshing list (datagrid of questions)")

  global MRA_QUESTIONS_LIST
  # wipe list in case we're reloading (this function should only be called once for initial load of MRA template)
  MRA_QUESTIONS_LIST = []

  #MessageBox.Show("Genating SQL...", "Refreshing list (datagrid of questions)")
  # firstly, we'll get main Question structure from source table (MRAv2_Templates)
  mySQL = """SELECT MRAT.QuestionGroup, MRAT.QuestionOrder, MRAT.QuestionID, MRAQ.QuestionText
             FROM Usr_MRAv2_Templates MRAT
                LEFT JOIN Usr_MRAv2_Question MRAQ ON MRAT.QuestionID = MRAQ.QuestionID
             WHERE MRAT.TemplateID = {0}
             GROUP BY MRAT.QuestionGroup, MRAT.QuestionOrder, MRAT.QuestionID, MRAQ.QuestionText
             ORDER BY MRAT.QuestionGroup, MRAT.QuestionOrder;""".format(lbl_MRA_TemplateID.Content)
  
  #MessageBox.Show("SQL: " + str(mySQL) + "\n\nRefreshing list (datagrid of questions)", "Debug: Populating List of Questions (Preview MRA)")

  _tikitDbAccess.Open(mySQL)
  dr = _tikitDbAccess._dr
  if dr is not None and dr.HasRows:
    while dr.Read():
      group_name = "" if dr.IsDBNull(0) else dr.GetValue(0)
      order_no = to_int(dr.GetValue(1))
      qid = to_int(dr.GetValue(2))
      qtext = "" if dr.IsDBNull(3) else dr.GetString(3)

      answers = MRA_ANSWERS_BY_QID.get(qid, [])
      MRA_QUESTIONS_LIST.append(MatterQuestionItem(group_name, order_no, qid, qtext, answers))

    dr.Close()
  _tikitDbAccess.Close()
  

  # then, we'll overlay existing selections for this matter (if any)
  MRA_load_MatterSelections_toMemory()
  MRA_apply_existing_selections()

  # and now we have a list of question items in memory for this template, with their available answers;
  # we can bind this to the datagrid and it should show the questions grouped by 'QuestionGroup' with a
  # combo box of available answers for each question (bound to the 'SelectedAnswerID' property of the
  # MatterQuestionItem, which will allow us to easily get the selected answer and its score/email comment (when
  # user selects an answer in the preview)
  # create observable collection for WPF and bind to datagrid; this should show the questions grouped by 'QuestionGroup' with a combo box of available answers for each question (bound to the 'SelectedAnswerID' property of the QuestionItem, which will allow us to easily get the selected answer and its score/email comment when user selects an answer in the preview)
  view = ListCollectionView(MRA_QUESTIONS_LIST)
  view.GroupDescriptions.Add(PropertyGroupDescription("QuestionGroup"))
  dg_MRA.ItemsSource = view

  has_items = (len(MRA_QUESTIONS_LIST) > 0)
  grid_MRA.Visibility = Visibility.Visible if has_items else Visibility.Collapsed
  tb_NoMRA_Qs.Visibility = Visibility.Collapsed if has_items else Visibility.Visible

  if has_items:
    # defer until UI has built group containers
    dg_MRA.Dispatcher.BeginInvoke(
                  DispatcherPriority.ContextIdle,
                  Action(_select_first_MRA_row)
                  )

  return

def MRA_load_MatterSelections_toMemory():
  #  Loads existing selections for this matter+mraID into MRA_MATTER_SELECTIONS_BY_QID:
  #  AnswerID + Comments per QuestionID
  global MRA_MATTER_SELECTIONS_BY_QID
  MRA_MATTER_SELECTIONS_BY_QID = {}

  # These should already be set before showing this tab:
  entity = str(_tikitEntity) 
  matter = to_int(_tikitMatter)
  mra_id = to_int(lbl_MRA_ID.Content)

  if not entity or matter is None or mra_id <= 0:
    # If you haven't wired these labels yet, just return silently
    return

  sql = """SELECT QuestionID, AnswerID, ISNULL(Comments,'')
           FROM Usr_MRAv2_MatterDetails
           WHERE EntityRef = '{0}' AND MatterNo = {1} AND mraID = {2}
           ORDER BY DisplayOrder;""".format(entity, to_int(matter), to_int(mra_id))

  _tikitDbAccess.Open(sql)
  dr = _tikitDbAccess._dr
  if dr is not None and dr.HasRows:
    while dr.Read():
      qid = to_int(dr.GetValue(0), 0)
      aid = to_int(dr.GetValue(1), 0)
      cmt = "" if dr.IsDBNull(2) else dr.GetString(2)

      if qid > 0:
        MRA_MATTER_SELECTIONS_BY_QID[qid] = {
          "AnswerID": (aid if aid > 0 else None),
          "Comments": cmt or ""
        }
    dr.Close()
  _tikitDbAccess.Close()
  return

def MRA_apply_existing_selections():
  # Overlays saved matter answers/comments onto the in-memory Question objects.
  # Must be called AFTER MRA_QUESTIONS_LIST is built and AvailableAnswers populated.
  global MRA_QUESTIONS_LIST

  if not MRA_QUESTIONS_LIST:
    return

  # quick lookup AnswerItem by (qid, aid)
  for q in MRA_QUESTIONS_LIST:
    sel = MRA_MATTER_SELECTIONS_BY_QID.get(to_int(q.QuestionID), None)
    if sel is None:
      continue

    # Comments
    try:
      q.UserComment = sel.get("Comments", "") or ""
    except:
      pass

    # Selected answer
    aid = sel.get("AnswerID", None)
    if aid is None:
      continue

    found = None
    try:
      for a in q.AvailableAnswers:
        if to_int(getattr(a, "AnswerID", None)) == to_int(aid):
          found = a
          break
    except:
      found = None

    # Important: set the object, not the ID
    if found is not None:
      try:
        q.SelectedAnswer = found
      except:
        pass

  return


def _select_first_MRA_row():
  if len(MRA_QUESTIONS_LIST) <= 0:
    return

  first = MRA_QUESTIONS_LIST[0]

  # force containers to exist
  try:
    dg_MRA.UpdateLayout()
  except:
    pass

  # 1) Force a real selection change event (important with Grouping DataGrids), select nothing and then first row
  dg_MRA.SelectedItem = None
  dg_MRA.SelectedItem = first

  # 2) ensure CurrentItem is aligned
  view = _get_preview_view()
  if view is not None:
    try:
      view.MoveCurrentTo(first)
    except:
      pass

  # 3) Commit "current cell" so WPF treats it as a real row selection
  try:
    if dg_MRA.Columns.Count > 0:
      dg_MRA.CurrentCell = DataGridCellInfo(first, dg_MRA.Columns[0])
  except:
    pass

  try:
    dg_MRA.ScrollIntoView(first)
    dg_MRA.Focus()
  except:
    pass

  _sync_combo_to_current_row()
  MRA_RecalcTotalScore()


def MRA_AdvanceToNextQuestion():
  # Uses the real list, so grouping doesn't break indices
  if len(MRA_QUESTIONS_LIST) <= 0:
    return

  curr = _get_current_question()
  if curr is None:
    # if something went odd, just go to first
    _select_first_MRA_row()
    return

  try:
    idx = MRA_QUESTIONS_LIST.index(curr)
  except:
    # CurrentItem might be a group wrapper; fall back to SelectedItem
    try:
      idx = MRA_QUESTIONS_LIST.index(dg_MRA.SelectedItem)
    except:
      _select_first_MRA_row()
      return

  next_idx = idx + 1
  if next_idx >= len(MRA_QUESTIONS_LIST):
    next_idx = 0  # wrap to first; change this if you'd rather stop at end

  nxt = MRA_QUESTIONS_LIST[next_idx]

  # Update selection + CurrentItem + scroll, then sync right panel combo
  dg_MRA.SelectedItem = nxt

  view = _get_preview_view()
  if view is not None:
    try:
      view.MoveCurrentTo(nxt)
    except:
      pass

  try:
    dg_MRA.ScrollIntoView(nxt)
  except:
    pass

  _sync_combo_to_current_row()
  return
  

def dg_MRA_SelectionChanged(s, event):
  # defer slightly so CurrentItem is updated, especially with grouping
  try:
    dg_MRA.Dispatcher.BeginInvoke(
      DispatcherPriority.Background,
      Action(_sync_combo_to_current_row)
    )
  except:
    _sync_combo_to_current_row()
  return


def _current_MRA_row():
  # Prefer CurrentItem (works properly with grouping)
  view = CollectionViewSource.GetDefaultView(dg_MRA.ItemsSource)
  if view is not None:
    try:
      return view.CurrentItem
    except:
      pass
  return dg_MRA.SelectedItem


def cbo_MRA_SelectedComboAnswer_SelectionChanged(s, event):
  global _preview_combo_syncing
  if _preview_combo_syncing:
    return  # ignore programmatic sync changes

  q = _get_current_question()
  if q is None or not hasattr(q, "SelectedAnswer"):
    return

  ans = cbo_MRA_SelectedComboAnswer.SelectedItem

  # IMPORTANT: ignore the transient "None" that occurs when ItemsSource swaps
  # unless the user genuinely cleared it (rare in your UI)
  if ans is None:
    return

  q.SelectedAnswer = ans

  view = _get_preview_view()
  if view is not None:
    view.Refresh()

  MRA_RecalcTotalScore()
  ##MessageBox.Show("Combo Selection Changed! Selected Question: " + str(row.QuestionText) + "\nSelected AnswerID: " + str(row.SelectedAnswerID) + "\nSelected Answer Text: " + str(row.SelectedAnswerText) + "\nSelected Answer Score: " + str(row.SelectedAnswerScore) + "\nSelected Answer Email Comment: " + str(row.SelectedAnswerEmailComment))
  return

def btn_MRASaveAnswer_Click(s, event):
  q = _current_MRA_row()
  if q is None:
    MessageBox.Show("No question selected to save answer for!", "Save")
    return

  a = getattr(q, "SelectedAnswer", None)
  if a is None:
    MessageBox.Show("No answer selected for the current question!", "Save")
    return

  #MessageBox.Show("Saved: QID={0}, AnswerID={1}".format(q.QuestionID, a.AnswerID))
  # if 'move to next' checkbox is ticked, then move to next question
  try:
    auto = (chk_MRA_AutoSelectNext.IsChecked == True)
  except:
    auto = False

  if auto:
    # Defer slightly so UI settles before we change selection
    try:
      dg_MRA.Dispatcher.BeginInvoke(
        DispatcherPriority.Background,
        lambda: MRA_AdvanceToNextQuestion()
      )
    except:
      MRA_AdvanceToNextQuestion()

  return

## Helper functions to sync the ComboBox to datagrid (in terms of displaying value)
def _get_preview_view():
  return CollectionViewSource.GetDefaultView(dg_MRA.ItemsSource)

def _get_current_question():
  view = _get_preview_view()
  if view is not None:
    try:
      return view.CurrentItem
    except:
      pass
  return dg_MRA.SelectedItem

def _sync_combo_to_current_row():
  global _preview_combo_syncing
  q = _get_current_question()
  if q is None or not hasattr(q, "AvailableAnswers"):
    return

  _preview_combo_syncing = True
  try:
    # Make sure combo is showing this row’s answers
    try:
      cbo_MRA_SelectedComboAnswer.ItemsSource = q.AvailableAnswers
    except:
      pass

    target = getattr(q, "SelectedAnswer", None)
    if target is None:
      cbo_MRA_SelectedComboAnswer.SelectedItem = None
      return

    # Choose matching item by AnswerID
    tid = getattr(target, "AnswerID", None)
    found = None
    for a in q.AvailableAnswers:
      if getattr(a, "AnswerID", None) == tid:
        found = a
        break

    cbo_MRA_SelectedComboAnswer.SelectedItem = found
  finally:
    _preview_combo_syncing = False


def MRA_RecalcTotalScore():
  # This function will Recalc the total score for selected answers and will also update the Total Questions / # answered
  # fields too

  # first, update totalQs and totalAnswered
  answered = 0
  for q in MRA_QUESTIONS_LIST:
    if getattr(q, "SelectedAnswer", None) is not None:
      answered += 1

  lbl_TotalQs.Content = str(len(MRA_QUESTIONS_LIST))
  lbl_TotalAnswered.Content = str(answered)

  # finally, calculate score and update label
  total = 0
  try:
    for q in MRA_QUESTIONS_LIST:
      # SelectedAnswerScore is 0 when no answer selected
      try:
        total += int(getattr(q, "SelectedAnswerScore", 0) or 0)
      except:
        pass
  except:
    total = 0

  lbl_MRA_Score.Content = str(total)

  # now work out 'category' based on score and thresholds; we have two thresholds: MediumFrom and HighFrom; if score is below MediumFrom, it's Low Risk; if it's between MediumFrom and HighFrom, it's Medium Risk; if it's above HighFrom, it's High Risk
  if total < to_int(lbl_ScoreTrigger_Medium.Content): 
    category = "Low"
    categoryNum = 1 
  elif total >= to_int(lbl_ScoreTrigger_Medium.Content) and total < to_int(lbl_ScoreTrigger_High.Content):
    category = "Medium" 
    categoryNum = 2
  elif total >= to_int(lbl_ScoreTrigger_High.Content):
    category = "High" 
    categoryNum = 3
  else: 
    category = "-"
    categoryNum = 0 
  
  lbl_MRA_RiskCategory.Content = category
  lbl_MRA_RiskCategoryID.Content = str(categoryNum)
  return total

################################################################

def btn_MRA_BackToOverview_Click(s, event):
  # This function should clear the 'MRA Questions' tab and take us back to the 'Overview' tab
  MRA_BackToOverview()
  return

def MRA_BackToOverview():
  # formally we were also setting the 'MinsToComplete' in 'Usr_MRA_Overview' table when user clicked 'Save as Draft' or 'Submit',
  # but as I didn't add field into new 'Usr_MRAv2_MatterHeader' table, we're not doing at the mo. 
  # If we were to re-add, I'd recommend we calculate 'seconds' rather than 'minutes', because it would be a lot more accurate (eg: 
  #  on the 'minutes' approach, it only counts whole minutes, so if took you 59 seconds, it would register 0 etc!)
  # I initially thought storing 'minutes' would be easier for reporting, however, due to accuracy issue, I think it's better to store 'seconds'
  #  and then just convert to minutes (rounding as needed) in any reports/dashboard etc. that we build on top.
  ti_Main.Visibility = Visibility.Visible
  ti_Main.IsSelected = True
  ti_MRA.Visibility = Visibility.Collapsed
  return


def flatten_matter_rows(store_unanswered=True):
  # Returns list[dict] ready for INSERT into Usr_MRAv2_MatterDetails.
  # Uses current on-screen order (MRA_QUESTIONS_LIST), so grouping won't break order.
  rows = []

  entity = str(_tikitEntity)  # 15-char
  matter = to_int(_tikitMatter)
  mra_id = to_int(lbl_MRA_ID.Content)

  if not entity or matter <= 0 or mra_id <= 0:
    raise Exception("Cannot save: missing Entity/Matter/mraID context.")

  display = 1
  for q in MRA_QUESTIONS_LIST:
    qid = to_int(getattr(q, "QuestionID", 0))
    if qid <= 0:
      continue

    group_name = getattr(q, "QuestionGroup", "") or ""
    ans_obj = getattr(q, "SelectedAnswer", None)
    aid = None if ans_obj is None else to_int(getattr(ans_obj, "AnswerID", None), 0)
    score = 0 if ans_obj is None else to_int(getattr(ans_obj, "Score", 0), 0)
    email_comment = "" if ans_obj is None else (getattr(ans_obj, "EmailComment", "") or "")
    comments = getattr(q, "UserComment", "") or ""

    if ans_obj is None and not store_unanswered:
      # skip unanswered entirely
      display += 1
      continue

    rows.append({
      "EntityRef": entity,
      "MatterNo": matter,
      "mraID": mra_id,
      "QuestionID": qid,
      "AnswerID": (aid if (aid is not None and aid > 0) else None),
      "Score": score,
      "Comments": comments,
      "DisplayOrder": display,
      "EmailComment": email_comment,
      "GroupName": group_name
    })

    display += 1

  return rows

def save_matterdetails_to_db(store_unanswered=True):
  entity = str(_tikitEntity)
  matter = to_int(_tikitMatter)
  mra_id = to_int(lbl_MRA_ID.Content)

  rows = flatten_matter_rows(store_unanswered=store_unanswered)

  # Delete existing matter detail rows for this MRA instance
  delete_sql = """DELETE FROM Usr_MRAv2_MatterDetails
                  WHERE EntityRef = '{0}' AND MatterNo = {1} AND mraID = {2};""".format(
                    sql_escape(entity), matter, mra_id
                  )

  sql_batch = [delete_sql]

  if rows:
    values_sql = []
    for r in rows:
      # AnswerID can be NULL (Draft/unanswered)
      aid_sql = "NULL" if r["AnswerID"] is None else str(to_int(r["AnswerID"]))
      values_sql.append(
        "('{EntityRef}', {MatterNo}, {mraID}, '{QuestionGroup}', {QuestionID}, {AnswerID}, {Score}, '{Comments}', {DisplayOrder}, '{EmailComment}')".format(
          EntityRef=sql_escape(r["EntityRef"]),
          MatterNo=to_int(r["MatterNo"]),
          mraID=to_int(r["mraID"]),
          QuestionGroup=sql_escape(r["GroupName"]),
          QuestionID=to_int(r["QuestionID"]),
          AnswerID=aid_sql,
          Score=to_int(r["Score"]),
          Comments=sql_escape(r["Comments"]),
          DisplayOrder=to_int(r["DisplayOrder"]),
          EmailComment=sql_escape(r["EmailComment"])
        )
      )

    insert_sql = (
      "INSERT INTO Usr_MRAv2_MatterDetails "
      "(EntityRef, MatterNo, mraID, QuestionGroup, QuestionID, AnswerID, Score, Comments, DisplayOrder, EmailComment) "
      "VALUES {0};".format(", ".join(values_sql))
    )
    sql_batch.append(insert_sql)

  # Execute as one batch (preferred)
  batch_sql = ";\r\n".join(sql_batch) + ";"
  runSQL(batch_sql)
  return


def btn_MRA_SaveAsDraft_Click(s, event):
  # This function will return back to the 'Overview' screen and save the MRA in a 'draft' (incomplete) state
  
  # save current answers to db:
  save_matterdetails_to_db(store_unanswered=False)

  # update 'Status'
  MRA_setStatus(lbl_MRA_ID.Content, 'Draft')
  
  # update the Risk Category (sets MRA level, and in Matter Properties)
  #! MRA_UpdateRiskCategory(s, event)
  
  # update main 'Risk Status' label on 'Overview' tab
  setMasterRiskStatus(s, event)
  # refresh main overview datagrid
  dg_MRAFR_Refresh(s, event)
  # go back to overview tab
  MRA_BackToOverview()
  return


def validate_all_answered():
  # This function checks that every question has an answer selected; if any are missing, it returns False
  # and a message listing the first few missing questions; if all are answered, it returns True and an empty message
  missing = []
  for q in MRA_QUESTIONS_LIST:
    if getattr(q, "SelectedAnswer", None) is None:
      missing.append(getattr(q, "QuestionText", ""))
  if missing:
    return (False, "Please answer all questions before submitting.\n\nFirst missing:\n• " + "\n• ".join(missing[:10]))
  return (True, "")

# helper functions to avoid duplication with MRA Submit
def is_high_risk():
  # returns True/False based on whether the current risk category is 'High' - this is used to determine whether we need to trigger the HOD approval process or not
  return str(lbl_MRA_RiskCategory.Content) == "High"

def current_user_is_fee_earner():
  # returns True/False based on whether the current user is the same as the Matter Fee Earner (based on tb_FERef.Text) - this is used to determine which approval rules to apply when submitting
  return str(_tikitUser) == str(tb_FERef.Text)

def can_auto_approve(high_risk):
  # Only relevant for high risk:
  # - if user is FE: canApproveSelf(user)
  # - else: canUserApproveFeeEarner(user, FE)
  
  if not high_risk:
    return False

  if current_user_is_fee_earner():
    return canApproveSelf(userToCheck=_tikitUser) == True
  else:
    return canUserApproveFeeEarner(UserToCheck=_tikitUser, FeeEarner=tb_FERef.Text) == True

def get_user_email(code):
  # returns the email address for a given user code; used for getting HOD email to send to when submitting high risk MRA
  return runSQL("SELECT EMailExternal FROM Users WHERE Code = '{0}'".format(code), False, '', '')

def get_user_name(code):
  # returns the forename (or full name if forename is null) for a given user code; 
  # used for getting HOD name to personalise email when submitting high risk MRA
  return runSQL("SELECT ISNULL(Forename, FullName) FROM Users WHERE Code = '{0}'".format(code), False, '', '')

def get_two_user_emails(code1, code2):
  # returns "a@x; b@y" (handles NULLs defensively)
  s = runSQL("SELECT STRING_AGG(EMailExternal, '; ') FROM Users WHERE Code IN ('{0}', '{1}')".format(code1, code2), False, '', '')
  return "" if s is None else str(s)

def request_hod_email_to_for_fee_earner(fee_earner_code, include_current_user=False):
  hod = getUsersApproversEmail(forUser=fee_earner_code) or ""
  if not include_current_user:
    return hod

  both = get_two_user_emails(_tikitUser, fee_earner_code)
  # combine both + hod
  # avoid double ;; if either side blank
  parts = [p.strip() for p in [both, hod] if p and str(p).strip()]
  return "; ".join(parts)

def decide_submit_outcome():
  # determines the 'outcome' of submitting the MRA, which will determine which TaskCentre email gets triggered; this is based on:
  # 1) whether the MRA is high risk or not (based on lbl_MRA_RiskCategory.Content)
  # 2) whether the current user is the Matter Fee Earner or not (based on tb_FERef.Text vs _tikitUser)
  # 3) whether the current user has approval rights (either canApproveSelf if FE, or canUserApproveFeeEarner if not FE)
  # The outcome will be one of:
  # - 'Submit_MRA_Standard' (not high risk, so no HOD approval needed; email goes to FE only)
  # - 'Submit_MRA_HighRisk_AutoApproved' (high risk but current user can approve, so no HOD approval needed; email goes to FE only)
  # - 'Submit_MRA_HighRisk_RequestHOD' (high risk and current user cannot approve, so HOD approval needed; email goes to HOD and FE)
  # - 'Submit_MRA_OnBehalf' (not high risk, but current user is not FE, so submitting on behalf of FE; email goes to FE only)
  high = is_high_risk()
  is_fe = current_user_is_fee_earner()

  if not high:
    # Standard risk
    return OUTCOME_SUBMIT_STD if is_fe else OUTCOME_ON_BEHALF

  # High risk
  if can_auto_approve(high_risk=True):
    return OUTCOME_AUTO_APPROVE

  # all else fallback to needing HOD approval
  return OUTCOME_REQUEST_HOD


def execute_submit_outcome(outcome, mraID, mraName, ourRef, matDesc, clName, feEmail, feName, riskRatingStr):

  if outcome == OUTCOME_AUTO_APPROVE:
    # Auto-approve triggers its own email via HOD_Approves_Item
    return HOD_Approves_Item(my_mraID=mraID, myEntRef=_tikitEntity, myMatNo=_tikitMatter, myMRADesc=mraName)

  if outcome == OUTCOME_REQUEST_HOD:
    # High risk, not authorised to auto approve -> request HOD
    if current_user_is_fee_earner():
      # Email HOD, CC FE (your existing behaviour)
      hodEmails = request_hod_email_to_for_fee_earner(fee_earner_code=_tikitUser, include_current_user=False)
      insert_into_MRAEvents(
        userRef=_tikitUser, triggerText='Submit_MRA_HighRisk', ov_ID=mraID,
        emailTo=hodEmails, emailCC=feEmail, toUserName=feName,
        ourRef=ourRef, matterDesc=matDesc, clientName=clName,
        addtl1=mraName, addtl2=riskRatingStr
      )
      return

    # Not FE: include current user + FE + HOD
    emailToAddr = request_hod_email_to_for_fee_earner(fee_earner_code=tb_FERef.Text, include_current_user=True)
    emailToName = get_user_name(code=_tikitUser)
    insert_into_MRAEvents(
      userRef=_tikitUser, triggerText='Submit_MRA_HighRisk', ov_ID=mraID,
      emailTo=emailToAddr, emailCC=feEmail, toUserName=emailToName,
      ourRef=ourRef, matterDesc=matDesc, clientName=clName,
      addtl1=mraName, addtl2=riskRatingStr
    )
    return

  if outcome == OUTCOME_SUBMIT_STD:
    # Standard risk + FE submitting
    insert_into_MRAEvents(
      userRef=_tikitUser, triggerText='Submit_MRA', ov_ID=mraID,
      emailTo=feEmail, emailCC='', toUserName=feName,
      ourRef=ourRef, matterDesc=matDesc, clientName=clName,
      addtl1=mraName, addtl2=riskRatingStr
    )
    return

  if outcome == OUTCOME_ON_BEHALF:
    # Standard risk + someone else submitting
    emailToAddr = get_user_email(code=_tikitUser)
    emailToName = get_user_name(code=_tikitUser)
    insert_into_MRAEvents(
      userRef=_tikitUser, triggerText='Submit_MRA_onBehalfOf', ov_ID=mraID,
      emailTo=emailToAddr, emailCC=feEmail, toUserName=emailToName,
      ourRef=ourRef, matterDesc=matDesc, clientName=clName,
      addtl1=mraName, addtl2=riskRatingStr
    )
    return

  raise Exception("Unknown submit outcome: " + str(outcome))


def btn_MRA_Submit_Click(s, event):
  # This function will return back to the 'Overview' screen and save the MRA in a 'Complete' state
  # (1: Saves All Q&A back to MRAv2_MatterDetails; 2: Updates status to 'Complete'/'Complete - Pending Approval' 
  #  in MRAv2_MatterHeader; 3: Triggers task centre emails based on rules)

  # first check all questions have an answer; if not, show message and exit without saving
  ok, msg = validate_all_answered()
  if not ok:
    MessageBox.Show(msg, "Cannot Submit", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    return

  # save details to DB
  save_matterdetails_to_db(store_unanswered=False)  # or True; your choice

  # get necessary details for email triggers (we have these on screen, so no need for extra SQL calls;
  # from the forms' Header details (groupbox)
  feEmail = tb_FE_Email.Text
  feName = tb_FE_Forename.Text
  matDesc = tb_MatterDesc.Text
  matDesc = matDesc.replace("'", "''")
  clName = tb_ClientName.Text
  clName = clName.replace("'", "''")
  mraName = lbl_MRA_Name.Content
  mraName = mraName.replace("'", "''")
  ourRef = tb_OurRef.Text
  # from the 'Edit MRA' tab
  mraID = lbl_MRA_ID.Content
  riskRating = lbl_MRA_RiskCategory.Content      # = Low | Medium | High
  riskRatingID = lbl_MRA_RiskCategoryID.Content  # = 1 | 2 | 3 (based on above mapping in MRA_RecalcTotalScore)

  # and execute the applicable path based on the rules (see helper functions above to determine outcome, and then execute that outcome)
  outcome = decide_submit_outcome()
  execute_submit_outcome(outcome, mraID, mraName, ourRef, matDesc, clName, feEmail, feName, riskRating)

  # need to udpate Status and Risk Category in both the Overview table and the Matter Properties table; we'll use the same function for both, which takes care of the different SQL needed for each
  updateHeaderSQL = """UPDATE Usr_MRAv2_MatterHeader SET RiskRating = {riskRating}, 
                              Status = 'Complete', SubmittedBy = '{subBy}', SubmittedDate = GETDATE() 
                       WHERE mraID = {mraID} AND EntityRef = '{entRef}' AND MatterNo = {matNo}
                    """.format(riskRating=riskRatingID, subBy=_tikitUser,
                               mraID=mraID, entRef=_tikitEntity, matNo=_tikitMatter)

  # execute header update
  runSQL(updateHeaderSQL)
  # note: no longer using separate 'MRA_setStatus' function here as we also need to update RiskRating at the same time, 
  # and it's more efficient to do in one SQL statement. The only other thing that function did was to put the new status
  # in the 'lbl_MRA_Status' label, but since we're going back to the overview immediately after, there's no need to update
  #  that label here (and risk category is already updated on the label in case we need it for email triggers) 

  # update main 'Risk Status' label on 'Overview' tab
  setMasterRiskStatus(s, event)
  # refresh main overview datagrid
  dg_MRAFR_Refresh(s, event)
  # go back to overview tab
  MRA_BackToOverview()
  return


def insert_into_MRAEvents(userRef, triggerText, ov_ID, emailTo, emailCC, toUserName, ourRef, matterDesc, clientName, addtl1, addtl2):
  # This function will insert the passed details into the Usr_MRA_Events table (which triggers Task Centre emails)

  tmpSQL = """INSERT INTO Usr_MRA_Events (Date, UserRef, ActionTrigger, OV_ID, EmailTo, EmailCC, ToUserName, OurRef, MatterDesc, 
              ClientName, Addtl1, Addtl2, FullEntityRef, Matter_No) 
              VALUES(GETDATE(), '{0}', '{1}', {2}, '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', '{11}', {12})""".format(userRef, triggerText, 
                      ov_ID, emailTo, emailCC, toUserName, ourRef, matterDesc, clientName, addtl1, addtl2, _tikitEntity, _tikitMatter)

  #runSQL(tmpSQL, True, "There was an error attempting to add a row to the Usr_MRA_Events table. \nConfirmation email may not be received\n\nSQL Used:\n{0}".format(tmpSQL), "ERROR: Attempting to save to 'Events' table...")
  try:
    _tikitResolver.Resolve("[SQL: {0}]".format(tmpSQL))
    return True
  except Exception as e:
    return False


def populate_MRA_DaysUntilLocked(expiryDate=None):
  # This function will populate the 'you only have x days to complete' message and controls whether it needs to be seen or not
  # Added a 'minus 1' to days following change to number of days)
  #MessageBox.Show("populate_MRA_DaysUntilLocked called with expiryDate={0}".format(expiryDate))

  # need to lookup current status (if complete, hide label and 'Save' buttons (and make 'back to overview' visible))
  if lbl_MRA_Status.Content == 'Draft':
    if expiryDate is not None:
      newExp = getSQLDate(expiryDate)
      #MessageBox.Show("Calculated newExp as {0}".format(newExp))
      daysTilLock = runSQL("SELECT DATEDIFF(DAY, GETDATE(), '{expDate}') - 1".format(expDate=newExp))
    else:
      daysTilLock = "N/A"
    tb_DaysUntilLocked.Text = "{0} day(s)".format(daysTilLock)
    tb_DaysUntilLocked.Visibility = Visibility.Visible
    tb_MatterWillBeLockedMsg.Visibility = Visibility.Visible
  else:
    tb_DaysUntilLocked.Visibility = Visibility.Collapsed
    tb_MatterWillBeLockedMsg.Visibility = Visibility.Collapsed
  return


 

# # # #   END OF:   M A T T E R   R I S K   A S S E S S M E N T    TAB   # # # #

# # # #   *F I L E   R E V I E W*    TAB    # # # #

class FR(object):
  def __init__(self, myID, myOrder, myQuestion, myAnswerText, myCorrActionID, myCANeeded, myCATaken, 
               myCAComplete, myQID, myGroup, myAllowNA, myCAtrigger, myHasOSCA, myAllowsComment, myComment):
 
    self.pvID = myID
    self.pvDO = myOrder
    self.pvQuestion = myQuestion
    self.QuestionID = myQID
    self.pvAnswerText = myAnswerText
    self.CorrActionID = myCorrActionID
    self.CorrActionNeeded = myCANeeded
    self.CorrActionTaken = myCATaken
    self.CorrActionComplete = myCAComplete
    self.QGroup = myGroup
    if myCAComplete == 0: 
      self.CorrActionCompleteTF = False
    else:
      self.CorrActionCompleteTF = True
    self.fr_AllowsNA = myAllowNA
    self.fr_CAtrigger = myCAtrigger
    self.fr_HasOSCA = myHasOSCA
    self.fr_AllowsComment = myAllowsComment
    self.fr_Comment = myComment
    return
    
  def __getitem__(self, index):
    if index == 'Order':
      return self.pvDO
    elif index == 'Question':
      return self.pvQuestion
    elif index == 'QuestionID':
      return self.QuestionID
    elif index == 'ID':
      return self.pvID
    elif index == 'AText':
      return self.pvAnswerText
    elif index == 'CorrActionID':
      return self.CorrActionID
    elif index == 'CorrActionNeeded':
      return self.CorrActionNeeded
    elif index == 'CorrActionTaken':
      return self.CorrActionTaken
    elif index == 'CorrActionComplete':
      return self.CorrActionComplete
    elif index == 'CorrActionCompleteTF':
      return self.CorrActionCompleteTF 
    elif index == 'QGroup':
      return self.QGroup
    elif index == 'AllowsNA':
      return self.fr_AllowsNA 
    elif index == 'CAtrigger':
      return self.fr_CAtrigger
    elif index == 'HasOSCA':
      return self.fr_HasOSCA
    elif index == 'AllowsComment':
      return self.fr_AllowsComment
    elif index == 'Comment':
      return self.fr_Comment
    else:
      return ''
      
def refresh_FR(s, event):
  # This function refreshes the File Review data grid (on the 'File Review' tab - where one Edits/adds answers to Qs)
  
  # generate the SQL to populate our datagraid
  # NB: MA.AuditPass = 0 means 'Fail'/False, 1 means 'Pass'/True (completed)
  mySQL = """SELECT '0-RowID' = FR.ID, '1-DisplayOrder' = FR.DisplayOrder, '2-Question Text' = TQ.QuestionText, 
			        '3-AnswerText' = FR.tbAnswerText, '4-CorrActionID' = FR.CorrActionID, '5-CANeeded' = MA.CorrActionNeeded, 
              '6-CATaken' = MA.CorrActionTaken, 
              '7-CA_PF' = MA.AuditPass, '8-Qid' = FR.QuestionID, '9-QGroup' = QG.Name, 
              '10-AllowsNA' = ISNULL(TQ.FR_Allow_NA_Answer, 'Y'), '11-CA Trigger' = ISNULL(TQ.FR_CorrAction_Trigger_Answer, 'No'), 
              '12-HasOSCA' = CASE WHEN MA.AuditPass = 0 AND FR.CorrActionID > 0 THEN 'Yes' ELSE 'No' END, 
              '13-AllowsComment' = ISNULL(TQ.FR_Allow_Comment, 'N'), '14-Comment' = FR.Notes 
            FROM Usr_MRA_Detail FR 
              LEFT OUTER JOIN Usr_MRA_TemplateQs TQ ON FR.QuestionID = TQ.ID 
              LEFT OUTER JOIN Matter_Audit MA ON FR.CorrActionID = MA.ID 
			        LEFT OUTER JOIN Usr_MRA_QGroups QG ON TQ.QGroupID = QG.ID
            WHERE FR.EntityRef = '{0}' AND FR.MatterNo = {1} AND FR.OV_ID = {2}
            ORDER BY QG.DisplayOrder, FR.DisplayOrder""".format(_tikitEntity, _tikitMatter, lbl_FR_ID.Content)

  _tikitDbAccess.Open(mySQL)
  myItems = []
  
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          iID = 0 if dr.IsDBNull(0) else dr.GetValue(0)
          iDO = 0 if dr.IsDBNull(1) else dr.GetValue(1)
          iQText = '-' if dr.IsDBNull(2) else dr.GetString(2)
          iAText = '' if dr.IsDBNull(3) else dr.GetString(3)
          iCAid = 0 if dr.IsDBNull(4) else dr.GetValue(4)
          iCANeeded = '' if dr.IsDBNull(5) else dr.GetString(5)
          iCATaken = '' if dr.IsDBNull(6) else dr.GetString(6)
          iCAComplete = 0 if dr.IsDBNull(7) else dr.GetValue(7)
          iQid = 0 if dr.IsDBNull(8) else dr.GetValue(8)
          iGroupName = '' if dr.IsDBNull(9) else dr.GetString(9)
          iAllowsNA = 'Y' if dr.IsDBNull(10) else dr.GetString(10)
          iCAtrigger = 'No' if dr.IsDBNull(11) else dr.GetString(11)
          iHasOSCA = 'No' if dr.IsDBNull(12) else dr.GetString(12)
          iAllowsComment = 'N' if dr.IsDBNull(13) else dr.GetString(13)
          iComment = '' if dr.IsDBNull(14) else dr.GetString(14)

          myItems.append(FR(myID=iID, myOrder=iDO, myQuestion=iQText, myAnswerText=iAText, myCorrActionID=iCAid, 
                            myCANeeded=iCANeeded, myCATaken=iCATaken, myCAComplete=iCAComplete, myQID=iQid, myGroup=iGroupName,
                            myAllowNA=iAllowsNA, myCAtrigger=iCAtrigger, myHasOSCA=iHasOSCA, myAllowsComment=iAllowsComment, myComment=iComment))
      
    dr.Close()
  _tikitDbAccess.Close()
  
  # as we have 'grouping' on the XAML (using the Expander control), here we need to create 'ListCollectionView'
  # in order to add 'GroupDescriptions' to properly 'group' our items accordingly (eg: uses the 'QGroup' name as a banner
  # above each question set)
  tmpC = ListCollectionView(myItems)
  tmpC.GroupDescriptions.Add(PropertyGroupDescription('QGroup'))
  dg_FR.ItemsSource = tmpC   
  
  # if nothing in the list
  if dg_FR.Items.Count == 0:
    # hide the datagrid and show the 'no questions' help label
    dg_FR.Visibility = Visibility.Hidden
    tb_NoFR_Qs.Visibility = Visibility.Visible
  else:
    # show the datagrid and hide the 'no questions' help label
    dg_FR.Visibility = Visibility.Visible
    tb_NoFR_Qs.Visibility = Visibility.Hidden
  return


def FR_optYes_Clicked(s, event):
  # This is the 'opt_Yes' button for File Reviews - clicking this will save 'Yes' and move to next question (if 'Auto move to next Q' ticked)
  # Linked to XAML control.event: opt_Yes.Click
  FR_SaveAnswer(s=s, event=event, answerValue = 'Yes')
  return


def FR_optNo_Clicked(s, event):
  # This is the 'opt_No' button for File Reviews - clicking this will save 'No' and move to next question (if 'Auto move to next Q' ticked)
  # Linked to XAML control.event: opt_No.Click
  FR_SaveAnswer(s=s, event=event, answerValue = 'No')
  return


def FR_optNA_Clicked(s, event):
  # This is the 'opt_NA' button for File Reviews - clicking this will save 'NA' and move to next question (if 'Auto move to next Q' ticked)
  # Linked to XAML control.event: opt_NA.Click
  FR_SaveAnswer(s=s, event=event, answerValue = 'N/A')
  return


def FR_SaveAnswer(s, event, answerValue = ''):
  # This is the main function to update the 'answer' on a File Review question
  # This is a generic function that each of the Answer option/radio buttons call and supply with a new 'answerValue'

  # get current values
  rowID = dg_FR.SelectedItem['ID']
  oldTextVal = dg_FR.SelectedItem['AText']
  questionID = dg_FR.SelectedItem['QuestionID']
  caID = dg_FR.SelectedItem['CorrActionID']
  CAtriggerText = lbl_FR_CAtrigger.Content

  caPreviouslyAdded = True if int(caID) > 0 else False

  # firstly check if there's any change to actually 'update' (if old and new value same, no need to update)
  if str(oldTextVal) == answerValue:
    # selected value is the same as it was previously, so do the 'auto advance' if checkbox ticked
    if chk_FR_AutoSelectNext.IsChecked == True:
      currDGindex = dg_FR.SelectedIndex
      FR_AutoAdvance(currDGindex, s, event)
    return
  
  # default 'caSQL' to nothing in case we don't set it in the 'if' traps below
  caSQL = ""
  newCA_ID = 0

  # We were previously adding a Corrective Action if the selected answer was 'No', however, since 22nd Jan 2025, 
  #  Amy would like to be able to select which answer triggers the CA, so we need to check the 'CA Trigger' field (temp stored in 'lbl_FR_CAtrigger')
  if answerValue == CAtriggerText:
  # OLD: if passed answer is no, then we need to add a new corrective action (if one not already added for this question)
  #if answerValue == 'No':
    if caPreviouslyAdded == False:
      # get the 'default' corrective action text (and replace any occurrence of 'dd/mm/yyyy' with the date 14 days from now)
      defaultCA = runSQL("SELECT ISNULL(REPLACE(FR_Default_Corrective_Action, 'dd/mm/yyyy', CONVERT(VARCHAR(12), DATEADD(day, 14, GETDATE()), 103)), '') FROM Usr_MRA_TemplateQs WHERE QuestionID = {0}".format(questionID))
      # put this into the 'ca' text box
      txt_CorractiveActionNeeded.Text = defaultCA
      # call function to create a new Corrective Action and return the ID of item added
      newCA_ID = add_CorrectiveAction(defaultCorrectiveActionText=defaultCA)
      lbl_CorrActionID.Content = str(newCA_ID)
      caSQL = ", CorrActionID = {0}".format(newCA_ID)
    # make sure 'Corrective Actions' area is visible
    stk_CorrectiveActions.Visibility = Visibility.Visible
  else:
    # hide the 'Corrective Actions' area as chosen answer isn't the 'Corrective Action' trigger
    stk_CorrectiveActions.Visibility = Visibility.Collapsed

  # form SQL to update with passed value
  updateSQL = "[SQL: UPDATE Usr_MRA_Detail SET tbAnswerText = '{0}', SelectedAnswerID = -2{1} WHERE ID = {2}]".format(answerValue, caSQL, rowID)
  
  try:
    _tikitResolver.Resolve(updateSQL)
    canContinue = True
  except:
    MessageBox.Show("There was an error updating the answer (no updates made!), using SQL:\n{0}".format(updateSQL), "Error: FR Preview - Updating Answer...")
    canContinue = False

  # if the user selects Yes after selecting no, we need to delete the already generated corrective action, so that it does not show up for the fee earner to action
  #if answerValue == 'Yes' and caPreviouslyAdded:
  # ^ UPDATED - no longer based on a a hard 'yes', checks against the 'CA Trigger' field (temp stored in 'lbl_FR_CAtrigger')
  if answerValue != CAtriggerText and caPreviouslyAdded:
    try:
      deletesql = "[SQL: DELETE FROM Matter_Audit WHERE EntityRef = '{0}' AND MatterNo = {1} AND ID = {2}]".format(_tikitEntity, _tikitMatter, caID)
      _tikitResolver.Resolve(deletesql) #Removes the row from the corrective actions table, to remove the CA
      updateSQL = "[SQL: UPDATE Usr_MRA_Detail SET CorrActionID = null WHERE ID = {0}]".format(rowID)
      _tikitResolver.Resolve(updateSQL) #updates the backing table for the datagrid to remove the corrective action ID, so that a new CA can be generated if yes selected again
      canContinue = True
    except:
      MessageBox.Show("Error deleting corrective action for this question.", "Error")
      canContinue = False
  
  # get current row index, and refresh DataGrid and finally select this Q again
  currDGindex = dg_FR.SelectedIndex
  refresh_FR(s, event)
  dg_FR.SelectedIndex = currDGindex
  # and scroll into view
  dg_FR.ScrollIntoView(dg_FR.Items[currDGindex])

  ## only 'auto-advance' if answer wasn't 'no' (eg: as user should complete corrective action)   ################################################################## MAY NEED TO DOUBLE-CHECK THIS WORKS
  if newCA_ID != 0: #str(newTextVal) == 'No' and newCA_ID != 0:
    txt_CorractiveActionNeeded.Focus()
    # remember to call function to make area visible if we do add a corrective action (and set 'True' or 'False' accordingly)
  else:
    # if we successfully updated something and 'auto select next' is ticked
    if canContinue == True and chk_FR_AutoSelectNext.IsChecked == True:
      #if chk_FR_AutoSelectNext.IsChecked == True:
      # select next Q
      FR_AutoAdvance(currDGindex, s, event)

  return


def FR_QComment_Save(s, event):
  #! Linked to XAML control.event: txt_FR_QComment.TextChanged/LostFocus
  # This function will save the 'Question Comment' text to the database when the text changes

  # get current values
  rowID = dg_FR.SelectedItem['ID']
  newTextVal = str(txt_FR_QComment.Text)
  newTextVal = newTextVal.replace("'", "''")
  updateSQL = "[SQL: UPDATE Usr_MRA_Detail SET Notes = '{0}' WHERE ID = {1}]".format(newTextVal, rowID)

  try:
    # run SQL to update text note
    _tikitResolver.Resolve(updateSQL)
    # get current selected index of datagrid, refresh list and re-select the current row
    currDGindex = dg_FR.SelectedIndex
    refresh_FR(s, event)
    FR_AutoAdvance(currDGindex, s, event)
  except:
    MessageBox.Show("There was an error updating the Comment (no updates made!), using SQL:\n" + updateSQL, "Error: File Review - Updating Comment...")
  return


def FR_AutoAdvance(currentDGindex, s, event):
  currPlusOne = currentDGindex + 1
  totalDGitems = dg_FR.Items.Count
  #MessageBox.Show("Current Index: {0}\nPlusOne: {1}\nTotalDGitems: {2}".format(currentDGindex, currPlusOne, totalDGitems), "Auto-Advance to next question...")
  
  # if current value plus one is equal to the total number of items in the datagrid
  if currPlusOne == totalDGitems:
    # select the first item
    dg_FR.SelectedIndex = 0
  else:
    # current value plus one is greater than or less than total items,  so set to next
    dg_FR.SelectedIndex = currPlusOne
  # finally scroll it into view
  dg_FR.ScrollIntoView(dg_FR.SelectedItem)
  return


def FR_SelectionChanged(s, event):
  # When the selection changes in the DataGrid dg_FR, this function will populate the hidden area with the selected question's details

  global token1
  myEntity = _tikitEntity
  myMatNo = _tikitMatter
  mraID = lbl_FR_ID.Content

  if dg_FR.SelectedIndex > -1:
    # something is selected in the datagrid, so populate our hidden labels with the selected values
    # (this enables us to test 'before' and 'after' on 'DataGrid.CellUpdate' event - and only update if value changed)
    lbl_FR_DGID.Content = dg_FR.SelectedItem['ID']
    lbl_FR_CurrVal.Content = dg_FR.SelectedItem['AText']
    lbl_CorrActionID.Content = dg_FR.SelectedItem['CorrActionID']
    lbl_CorrAct_QText.Text = dg_FR.SelectedItem['Question']
    lbl_FR_CAtrigger.Content = dg_FR.SelectedItem['CAtrigger']
    chk_FR_AllowsNA.IsChecked = False if dg_FR.SelectedItem['AllowsNA'] == 'N' else True
    chk_FR_AllowsNotes.IsChecked = False if dg_FR.SelectedItem['AllowsComment'] == 'N' else True
    txt_FR_QComment.Text = dg_FR.SelectedItem['Comment']
    txt_CorractiveActionNeeded.Text = dg_FR.SelectedItem['CorrActionNeeded']
    txt_CorractiveActionTaken.Text = dg_FR.SelectedItem['CorrActionTaken']
    chk_CorrectiveActionPassed.IsChecked = dg_FR.SelectedItem['CorrActionCompleteTF']
    
    # select/tick appropriate radio button
    opt_Yes.IsChecked = True if lbl_FR_CurrVal.Content == 'Yes' else False
    opt_No.IsChecked = True if lbl_FR_CurrVal.Content == 'No' else False
    opt_NA.IsChecked = True if lbl_FR_CurrVal.Content == 'N/A' else False
      
    # if there is a Corrective Action ID, then we need to show the 'Corrective Actions' area, otherwise, hide it.
    # NB: this CAN be done with just XAML (binding the 'Visibility' to the lbl_CorrActionID value)
    if dg_FR.SelectedItem['CorrActionID'] == 0:
      #stk_CorrectiveActions.IsEnabled = False
      stk_CorrectiveActions.Visibility = Visibility.Collapsed
    else:
      # set visibility of 'Corrective Actions' controls to true as there IS a CA
      stk_CorrectiveActions.Visibility = Visibility.Visible
      # also only enable controls if in edit mode (linking directly to option button so no manual 'if' needed)
      stk_CorrectiveActions.IsEnabled = opt_EditModeFR.IsChecked

  else:
    lbl_FR_DGID.Content = ''
    lbl_FR_CurrVal.Content = ''
    lbl_CorrActionID.Content = ''
    lbl_CorrAct_QText.Text =''
    txt_CorractiveActionNeeded.Text = ''
    txt_CorractiveActionTaken.Text = ''
    chk_CorrectiveActionPassed.IsChecked = False
    #stk_CorrectiveActions.IsEnabled = False
    stk_CorrectiveActions.Visibility = Visibility.Collapsed

    chk_FR_AllowsNA.IsChecked = False
    lbl_FR_CAtrigger.Content = ''
    opt_Yes.IsChecked = False
    opt_No.IsChecked = False
    opt_NA.IsChecked = False

  # MP: new 01/04/2025: Adding call to update 'header' stats
  if update_FR_Stats(ov_ID=mraID) == True and token1 == 0:
    token1 = 1
    result = MessageBox.Show("You have completed all questions and there are no Corrective Actions, do you want to mark this File Review complete?", "Confirmation", MessageBoxButtons.YesNo)
    if result == DialogResult.Yes:
      btn_FR_Submit_Click(s, event)
  return

  ## Louis added below 'auto-check' if any Q's still awaiting an answer, and if none, and no CA's, ask to 'complete'
  ## One issue though, it that this doesn't check the current status of the FR (from Overview table) to see if it's already marked as 'complete'
  ## get current status:
  #frStatus = runSQL("SELECT ISNULL(Status, '') FROM Usr_MRA_Overview WHERE ID = {0}".format(mraID))
  #
  ## get count of Questions without an Answer (eg: AnswerText = '')
  #countOfQsNoAnswer_s = """SELECT COUNT(QuestionID) FROM Usr_MRA_Detail MRAD
  #                         WHERE MRAD.EntityRef = '{0}' AND MRAD.MatterNo = {1} AND ISNULL(tbAnswerText, '') = '' 
  #                         AND MRAD.OV_ID = {2}""".format(myEntity, myMatNo, mraID)
  #
  #countOfQsNoAnswer_s = runSQL(countOfQsNoAnswer_s)
  #
  ## get count of Corrective Actions that are not yet complete (eg: MA.AuditPass = 0)... NB: this appears to also be looking for empty CA 'taken' text too...
  #OutstandingCAsCount = """SELECT COUNT(QuestionID) FROM Usr_MRA_Detail MRAD 
  #                              LEFT OUTER JOIN Matter_Audit MA ON MRAD.CorrActionID = MA.ID WHERE MA.AuditPass = 0 
  #                              AND MA.EntityRef = '{0}' AND MA.MatterNo = {1}
  #                              AND MA.CorrActionTaken = '' """.format(_tikitEntity, _tikitMatter) 
  #
  #OutstandingCAsCount = runSQL(OutstandingCAsCount)
  #
  ## now for logic: if no questions without an answer and no outstanding CAs, and we haven't already asked to mark as 'complete', ask if they want to
  #if int(float(countOfQsNoAnswer_s)) == 0 and token1 == 0 and int(float(OutstandingCAsCount)) == 0 and frStatus != 'Complete':
  #  token1 = 1
  #  result = MessageBox.Show("You have completed all questions and there are no Corrective Actions, do you want to mark this File Review complete?", "Confirmation", 
  #                      MessageBoxButtons.YesNo)
  #  if result == DialogResult.Yes:
  #    btn_FR_Submit_Click(s, event)
  #  else:
  #    return
  #
  #return


def FR_BackToOverview(s, event):
  # This function should clear the 'FR Questions' tab and take us back to the 'Overview' tab
  # refresh main overview datagrid
  dg_MRAFR_Refresh(s, event)
  dgCA_Overview_Refresh(s, event)
  
  # hide FR Questions tab and select overview
  ti_Main.Visibility = Visibility.Visible
  ti_Main.IsSelected = True
  ti_FR.Visibility = Visibility.Collapsed
  return


def btn_FR_Submit_Click(s, event):
  # This function will mark the current File Review as 'Complete' and trigger the Task Centre task to email the FE with any corrective actions
  #! Linked to XAML control.event: btn_FR_Submit.Click

  # get initial variables
  xovID = lbl_FR_ID.Content

   # call our new function to update stats on XAML
  update_FR_Stats(ov_ID=xovID)

  #countOfIncompleteCAs = int(tb_TotalOSCAs_FR.Text)
  countOfQuestions = int(tb_TotalQs_FR.Text)
  countAnswered = int(tb_TotalAnswered_FR.Text)
  countOfQsNoAnswer = int(countOfQuestions) - int(countAnswered)

  #countOfQuestions = runSQL("SELECT COUNT(QuestionID) FROM Usr_MRA_Detail WHERE OV_ID = {0}".format(ovID))
  #countOfQsNoAnswer = runSQL("SELECT COUNT(QuestionID) FROM Usr_MRA_Detail WHERE OV_ID = {0} AND ISNULL(tbAnswerText, '') = ''".format(ovID))

  # Firstly, check to see if all questions have an answer (if no, cannot mark as 'complete')
  if int(countOfQuestions) > 0:
    # set the 'temp' message to show if there are any questions missing an answer
    if int(countOfQsNoAnswer) == 1:
      tmpMsg = "Cannot mark this File Review as complete, because there is one question that doesn't yet have an answer!"
    if int(countOfQsNoAnswer) > 1:
      tmpMsg = "Cannot mark this File Review as complete, because there are {0} questions that haven't yet been answered!".format(countOfQsNoAnswer)
    
    # if count of quesions without an answer is greater than zero - alert user and exit function
    if int(countOfQsNoAnswer) > 0:
      MessageBox.Show(tmpMsg, "'Save and Send to Fee Earner' error...")
      return
  #else:
    # there doesn't appear to be any questions so what to do here??
  # else we continue on...
  #MessageBox.Show("ovID: {0}\nCoutOfQs: {1}\nCountOfQs_NoAnswer: {2}".format(ovID, countOfQuestions, countOfQsNoAnswer), "DEBUGGING - btn_FR_Submit_Click")

  # get amount of mins spent previously (this used to be inside following 'if', but pulled out to re-use for 'else' part)
  existingMins = runSQL("SELECT ISNULL(MinsToComplete, 0) FROM Usr_MRA_Overview WHERE ID = {0}".format(xovID))
  #MessageBox.Show("ovID: {0}\nCoutOfQs: {1}\nCountOfQs_NoAnswer: {2}\nexistingMins: {3}\nlbl_TimeEnteredFR.Content: {4}".format(ovID, countOfQuestions, countOfQsNoAnswer, existingMins, lbl_TimeEnteredFR.Content), "DEBUGGING - btn_FR_Submit_Click")

  # get Time to Complete
  if len(str(lbl_TimeEnteredFR.Content)) > 0:
    # get the time we entered the screen
    timeEntered = lbl_TimeEnteredFR.Content
    # calculate mins for THIS session (date diff between now and time started)
    thisSessionMins = runSQL("SELECT ISNULL(DATEDIFF(minute, '{0}', GETDATE()), 0)".format(timeEntered))
    # calculate new total time spent (old/existing mins plus this session)
    newTotalSpent = int(existingMins) + int(thisSessionMins)
  else:
    # as time doesn't appear to have been entered, we'll set 'new' time to 'old' time to avoid setting back to zero
    newTotalSpent = int(existingMins)

  #MessageBox.Show("existingMins: {0}\ntimeEntered: {1}\nnewTotalSpent: {2}".format(existingMins, timeEntered, newTotalSpent), "DEBUGGING - btn_FR_Submit_Click")
  # update main 'Overview' table with new 'Status'
  runSQL("UPDATE Usr_MRA_Overview SET SubmittedBy = '{0}', SubmittedDate = GETDATE(), MinsToComplete = {1} WHERE ID = {2}".format(_tikitUser, newTotalSpent, xovID))

  # now call function to set 'Status', and send email email from Task Centre to Fee Earner
  FR_checkForOSca_andFinalise(sender=s, e=event, ovID=int(xovID), callingFrom='btn_FR_Submit_Click')

  # refresh Corrective Actions datagrid on main page:
  dgCA_Overview_Refresh(s, event)

  # we've already refreshed the 'dg_MRAFR' datagrid (in above function), so now go back to main overview tab
  ti_Main.Visibility = Visibility.Visible
  ti_Main.IsSelected = True
  ti_FR.Visibility = Visibility.Collapsed
  return


def FR_UpdateReviewerWithActionTaken_Click(s, event)  :
  # This function will trigger a Task Centre task (Name: ''), which sends an email to the File Reviewer to advise that FE has
  # completed corrective actions, and needs to verify.
  #! Linked to XAML control.event: btn_FR_UpdateReviewerWithActionTaken.Click

  tmpTriggerText = 'FR_CorrectiveActions_Complete'

  # now set variables to pass into 'mra_events' table
  myEntity = str(_tikitEntity)
  myMatNo = _tikitMatter
  tmpOurRef = "{0}{1}/{2}".format(myEntity[0:3], myEntity[11:15], myMatNo)
  #### need to finish below variables
  ovID = runSQL(codeToRun="SELECT TOP 1 ID FROM Usr_MRA_Overview WHERE EntityRef = '{0}' AND MatterNo = {1} AND Status = 'With FE' ORDER BY DateAdded Desc".format(myEntity, myMatNo))
  tmpToUserName = runSQL(codeToRun="SELECT ISNULL(Forename, FullName) FROM Users WHERE Code = (SELECT FR_Reviewer FROM Usr_MRA_Overview WHERE ID = {0})".format(ovID), apostropheHandle=1)
  tmpMatDesc = runSQL(codeToRun="SELECT Description FROM Matters WHERE EntityRef = '{0}' AND Number = {1}".format(myEntity, myMatNo), apostropheHandle=1)
  tmpClName = runSQL(codeToRun="SELECT LegalName FROM Entities WHERE Code = '{0}'".format(myEntity), apostropheHandle=1)
  # email to = File Reviewer | email CC = current user
  tmpEmailTo = runSQL("SELECT EMailExternal FROM Users WHERE Code = (SELECT FR_Reviewer FROM Usr_MRA_Overview WHERE ID = {0})".format(ovID))
  tmpEmailCC = runSQL("SELECT EMailExternal FROM Users WHERE Code = '{0}'".format(_tikitUser))
  tmpLocalName =  runSQL("SELECT ISNULL(LocalName, 'File Review') FROM Usr_MRA_Overview WHERE ID = {0}".format(ovID))
  
  # form SQL to get the count of incomplete Corrective Actions for matter, and run
  countOfIncompleteCAs_SQL = """SELECT COUNT(QuestionID) FROM Usr_MRA_Detail MRAD 
                                LEFT OUTER JOIN Matter_Audit MA ON MRAD.CorrActionID = MA.ID WHERE MA.AuditPass = 0 
                                AND MA.EntityRef = '{0}' AND MA.MatterNo = {1}""".format(myEntity, myMatNo)
  countOfIncompleteCAs = runSQL(countOfIncompleteCAs_SQL)

  # Insert a record into MRA Events table to trigger email to FE
  insert_into_MRAEvents(userRef=_tikitUser, triggerText=tmpTriggerText, ov_ID=ovID, 
                        emailTo=tmpEmailTo, emailCC=tmpEmailCC, toUserName=tmpToUserName, 
                        ourRef=tmpOurRef, matterDesc=tmpMatDesc, clientName=tmpClName, 
                        addtl1=tmpLocalName, addtl2=countOfIncompleteCAs)
  MessageBox.Show("An email has been sent to the File Reviewer, with details of the Corrective Actions taken")
  return


def update_FR_Stats(ov_ID='ID'):
  # This function will update the stats shown in the 'header' area at top of the 'File Review' tab.

  # get Overview_ID
  if ov_ID == 'ID' or ov_ID == 0:
    # if no ID passed, get the ID from the label on the screen
    ov_ID = lbl_FR_ID.Content
  if ov_ID == 'ID':
    # if OV_ID still saying 'ID', then alert user and quit
    MessageBox("Overview ID doesn't appear to have been set!\nCannot update the header stats!", "Error: update_FR_Stats")
    return

  if int(ov_ID) > 0:
    # get the stats from the database
    totalQs = runSQL("SELECT COUNT(QuestionID) FROM Usr_MRA_Detail WHERE OV_ID = {0}".format(ov_ID))
    totalAnswered = runSQL("SELECT COUNT(QuestionID) FROM Usr_MRA_Detail WHERE OV_ID = {0} AND ISNULL(tbAnswerText, '') <> ''".format(ov_ID))
    totalCAs = runSQL("SELECT COUNT(QuestionID) FROM Usr_MRA_Detail MRAD LEFT OUTER JOIN Matter_Audit MA ON MRAD.CorrActionID = MA.ID WHERE ISNULL(MRAD.CorrActionID, '') <> '' AND MRAD.OV_ID = {0}".format(ov_ID))
    totalOSCAs = runSQL("SELECT COUNT(QuestionID) FROM Usr_MRA_Detail MRAD LEFT OUTER JOIN Matter_Audit MA ON MRAD.CorrActionID = MA.ID WHERE MA.AuditPass = 0 AND MRAD.OV_ID = {0}".format(ov_ID))
    # get the status of the File Review (from Overview table)
    status = runSQL("SELECT ISNULL(Status, 'Draft') FROM Usr_MRA_Overview WHERE ID = {0}".format(ov_ID))
  else:
    # if no ID, set all values to nothing
    status = 'Draft'
    totalQs = 0
    totalAnswered = 0
    totalCAs = 0
    totalOSCAs = 0

  # put values into fields on the XAML
  txt_FR_Status.Text = status
  tb_TotalQs_FR.Text = str(totalQs)
  tb_TotalAnswered_FR.Text = str(totalAnswered)
  tb_TotalCAs_FR.Text = str(totalCAs)
  tb_TotalOSCAs_FR.Text = str(totalOSCAs)

  # calc the number of Q's without an answer...
  QsWithNoAnswer = int(totalQs) - int(totalAnswered)

  # if no OS Q's, and no OS CAs, and Status isn't 'Complete', then return 'True' (so we can ask to mark as complete)
  if int(QsWithNoAnswer) == 0 and int(totalOSCAs) == 0 and status != 'Complete':
    return True
  else:
    return False

# # # #   END OF:   *F I L E    R E V I E W*    TAB   # # # # 

# # # #   C O R R E C T I V E   A C T I O N S   (ON FILE REVIEW TAB)  # # # #
  
def txt_CorractiveActionTaken_LostFocus(s, event):
  # This will commit any update to the database
  #! Linked to XAML control.event: txt_CorractiveActionTaken.LostFocus

  caTaken = txt_CorractiveActionTaken.Text
  caTaken = caTaken.replace("'", "''")
  
  update_SQL = "[SQL: UPDATE Matter_Audit SET CorrActionTaken = '{0}' WHERE ID = {1}]".format(caTaken, lbl_CorrActionID.Content)
  try:
    _tikitResolver.Resolve(update_SQL)
    tmpDGindex = dg_FR.SelectedIndex
    refresh_FR(s, event)
    dg_FR.SelectedIndex = tmpDGindex
    # and scroll it into view
    dg_FR.ScrollIntoView(dg_FR.Items[tmpDGindex])
  except:
    MessageBox.Show("There was an error saving the Corrective Action.\nUsing SQL:\n{0}".format(update_SQL), "Error: Saving Corrective Action...")
  return

def txt_CorractiveActionNeeded_LostFocus(s, event):
  # This will commit any update to the database
  #! Linked to XAML control.event: txt_CorractiveActionNeeded.LostFocus

  caNeeded = txt_CorractiveActionNeeded.Text
  caNeeded = caNeeded.replace("'", "''")
   
  update_SQL = "[SQL: UPDATE Matter_Audit SET CorrActionNeeded = '{0}' WHERE ID = {1}]".format(caNeeded, lbl_CorrActionID.Content)
  try:
    _tikitResolver.Resolve(update_SQL)
    tmpDGindex = dg_FR.SelectedIndex
    refresh_FR(s, event)
    dg_FR.SelectedIndex = tmpDGindex
    # and scroll it into view
    dg_FR.ScrollIntoView(dg_FR.Items[tmpDGindex])
  except:
    MessageBox.Show("There was an error saving the Corrective Action.\nUsing SQL:\n{0}".format(update_SQL), "Error: Saving Corrective Action...")
  return


def chk_CorrectiveActionPassed_Click(s, event):
  # This will commit any update to the database
  #! Linked to XAML control.event: chk_CorrectiveActionPassed.Click

  caComplete = 1 if chk_CorrectiveActionPassed.IsChecked == True else 0
  
  update_SQL = "[SQL: UPDATE Matter_Audit SET AuditPass = {0} WHERE ID = {1}]".format(caComplete, lbl_CorrActionID.Content)
  try:
    _tikitResolver.Resolve(update_SQL)
    tmpDGindex = dg_FR.SelectedIndex
    refresh_FR(s, event)
    dg_FR.SelectedIndex = tmpDGindex
    # and scroll it into view
    dg_FR.ScrollIntoView(dg_FR.Items[tmpDGindex])
  except:
    MessageBox.Show("There was an error saving the Corrective Action.\nUsing SQL:\n{0}".format(update_SQL), "Error: Saving Corrective Action...")
    return

  # new: check if any outstanding CA and if not, mark FR as complete
  FR_checkForOSca_andFinalise(sender=s, e=event, ovID=int(lbl_FR_ID.Content))
  # finally refresh the Corrective Actions datagrid and if 'ViewAll' selected, find item and re-select it
  dgCA_Overview_Refresh(s, event)
  return


def btn_CorrectiveAction_Save_Clicked(s, event):
  # This function will SAVE the details to the Matters Audit table (where Corrective Actions are stored)
  # NB: we don't need to 'insert' here, because we auto-add one upon user selecting 'No' answer (in FR_CellEditEnding function)  
  # Linked to XAML control.event: btn_CorrAction_Save.Click
  
  caNeeded = txt_CorractiveActionNeeded.Text
  caNeeded = caNeeded.replace("'", "''")
  caTaken = txt_CorractiveActionTaken.Text
  caTaken = caTaken.replace("'", "''")
  caComplete = 1 if chk_CorrectiveActionPassed.IsChecked == True else 0
  
  update_SQL = """[SQL: UPDATE Matter_Audit SET CorrActionNeeded = '{0}', CorrActionTaken = '{1}', AuditPass = {2} 
                        WHERE ID = {3}]""".format(caNeeded, caTaken, caComplete, lbl_CorrActionID.Content)
  try:
    _tikitResolver.Resolve(update_SQL)
    tmpDGindex = dg_FR.SelectedIndex
    refresh_FR(s, event)
    dg_FR.SelectedIndex = tmpDGindex
    # and scroll it into view
    dg_FR.ScrollIntoView(dg_FR.Items[tmpDGindex])
  except:
    MessageBox.Show("There was an error saving the Corrective Action.\nUsing SQL:\n{0}".format(update_SQL), "Error: Saving Corrective Action...")
    return

  # new: check if any outstanding CA and if not, mark FR as complete
  FR_checkForOSca_andFinalise(sender=s, e=event, ovID=int(lbl_FR_ID.Content))

  # get current DataGrid index position (for passing into 'FR_AutoAdvance' function)
  currDGindex = dg_FR.SelectedIndex
  # refresh the Corrective Actions datagrid 
  dgCA_Overview_Refresh(s, event)
  if chk_FR_AutoSelectNext.IsChecked == True:
    FR_AutoAdvance(currDGindex, s, event)
  return


def add_CorrectiveAction(defaultCorrectiveActionText = ''):
  # This function will add a new Corrective Action (if one doesn't already exist)
  
  # first, get details to allow us to generate SQL
  mFEref = tb_FERef.Text         #_tikitResolver.Resolve("[SQL: SELECT FeeEarnerRef FROM Matters WHERE EntityRef = '" + _tikitEntity + "' AND Number = " + str(_tikitMatter) + "]")
  # NB: I'm now wondering if 'Fee Earner' as shown on the Matter Audit screen in P4W is actually meant to show 'Reviewer' (person conducting File Review / adding Corrective Action)
  #     As Fee Earner already noted on Matter, it doesn't make sense that they need to be selected again on this screen (unless it's just lazyness on Advanced's part, with regard to reports)
  #     But... maybe it does need to be FE beause of report, in which case we may want another field to store 'Reviewer' in our 'Overview' table
  auditTypeID = _tikitResolver.Resolve("[SQL: SELECT AuditTypeID FROM Audit_Types WHERE Description = 'File Review']")
  
  caNeeded = defaultCorrectiveActionText     #txt_CorractiveActionNeeded.Text #add the sql to get the email comment here
  caNeeded = caNeeded.replace("'", "''")
  caTaken = str(txt_CorractiveActionTaken.Text)
  caTaken = caTaken.replace("'", "''")  
  
  # Code to insert a new Corrective Action
  newCA_SQL = """[SQL: INSERT INTO Matter_Audit (EntityRef, MatterNo, AuditPass, AuditDate, FeeEarnerRef, CorrActionNeeded, 
                                                 CorrActionTaken, AuditTypeRef, NextAuditDate) 
                       VALUES ('{0}', {1}, 0, GETDATE(), '{2}', '{3}', '{4}', {5}, DATEADD(DAY, 14, GETDATE()))]""".format(_tikitEntity, _tikitMatter, mFEref, caNeeded, caTaken, auditTypeID)

  
  try:
    _tikitResolver.Resolve(newCA_SQL)
  except:
    MessageBox.Show("There was an error adding a Corrective Action, using SQL:\n{0}".format(newCA_SQL), "Error: Adding new Corrective Action...")
  
  # try and get ID to return back to calling function
  newID_SQL = """[SQL: SELECT TOP 1 ID FROM Matter_Audit WHERE EntityRef = '{0}' AND MatterNo = {1} AND AuditTypeRef = {2} AND CorrActionNeeded = '{3}' 
                       AND CorrActionTaken = '{4}' ORDER BY NextAuditDate DESC]""".format(_tikitEntity, _tikitMatter, auditTypeID, caNeeded, caTaken)
  
  try:
    newID = _tikitResolver.Resolve(newID_SQL)
  except:
    newID = "0"
    MessageBox.Show("There was an error getting the ID of the newly added Corrective Action, using SQL:\n{0}".format(newID_SQL), "Error: Getting ID of newly added Corrective Action...")
  
  return int(newID)
  

# # # #   END OF:   C O R R E C T I V E   A C T I O N S   # # # #    


# # # # #   G E N E R I C   F U N C T I O N S   # # # # #
def runSQL(codeToRun, showError = False, errorMsgText = '', errorMsgTitle = '', apostropheHandle = 0, 
           useAlternativeResolver = False, returnType = 'Int'):
  # This function is written to handle and check inputted SQL code, and will return the result of the SQL code.
  # It first checks the length and wrapping of the code, then attempts to execute the SQL, it has an option apostrophe handler.
  # codeToRun     = Full SQL of code to run. No need to wrap in '[SQL: code_Here]' as we can do that here
  # showError     = True / False. Indicates whether or not to display message upon error
  # errorMsgText  = Text to display in the body of the message box upon error (note: actual SQL will automatically be included, so no need to re-supply that)
  # errorMsgTitle = Text to display in the title bar of the message box upon error
  # apostropheHandle = Toggle to escape apostrophes for the returned values
  # useAlternativeResolver = Toggle to use an alternative resolver for the SQL execution

  # Note: calling procedure can use like we do with '_tikitResolver()', that is: 
  # - tmpValue = runSQL("SELECT YEAR()", False, '', '')   # to capture value into a variable, or:
  # - runSQL("INSERT INTO x () VALUES()", False, '', '')  # to just run the SQL without saving to variable
  
  # if no code actually supplied, exit early
  if len(codeToRun) < 10:
    MessageBox.Show("The supplied 'codeToRun' doesn't appear long enough, please check and update this code if necessary.\nPassed SQL: " + str(codeToRun), "ERROR: runSQL...")
    return
  
  # Add '[SQL: ]' wrapper if not already included
  if codeToRun[:5] == "[SQL:":
    fCodeToRun = codeToRun
  else:
    fCodeToRun = "[SQL: {0}]".format(codeToRun)
  
  if useAlternativeResolver == False:
    # try to execute the SQL using standard 'tikitResolver.Resolve'
    try:
      tmpValue = _tikitResolver.Resolve(fCodeToRun)
      if apostropheHandle == 1:
        tmpValue = tmpValue.replace("'", "''")
      return tmpValue
    except:
      # there was an error... check to see if opted to show message or not...
      if showError == True:
        MessageBox.Show(str(errorMsgText) + "\nSQL used:\n" + str(codeToRun), errorMsgTitle)
      return ''
  else:
    # use longer method
    try:
      _tikitDbAccess.Open(codeToRun)
      if _tikitDbAccess._dr is not None:
        dr = _tikitDbAccess._dr
        if dr.HasRows:
          while dr.Read():

            if returnType == 'String':
              # return the first column as a string
              tmpValue = dr.GetString(0) if not dr.IsDBNull(0) else ''

            elif returnType == 'Int':
              # return the first column as an integer
              tmpValue = dr.GetValue(0) if not dr.IsDBNull(0) else 0
              
            #elif returnType == 'DataReader':
            #  # return the DataReader object itself
            #  return dr
        else:
          if returnType == 'String':
            # if no rows returned, return an empty string
            tmpValue = ''
          elif returnType == 'Int':
            # if no rows returned, return 0
            tmpValue = 0
        dr.Close()
      _tikitDbAccess.Close()

      return tmpValue
  
    except Exception as e:
      # if there was an error with the CTE supplied, we'll get to here, so update outputs accordingly
      if showError == True:
        MessageBox.Show("{0}\nSQL used:\n{1}\nException:{2}".format(errorMsgText, codeToRun, str(e)), errorMsgTitle)

      if returnType == 'Int':
        return 0
      else:
        return "Error"


    

def isUserAnApprovalUser(userToCheck):
  # This is a new function to replace the 'isActiveUserHOD()' function (from 7th August 2024) as we have now created an 'WhoApprovesMe' 
  # field in a new 'Usr_Approvals' table (user-level), that is better to check against.
  # 17th March 2025: Updated to point to new 'Usr_HODapprovals' table - a self-service table that allows HODs to choose their FEs
  
  ## old version using the 'Who Approves Me'
  #tmpCountAppearancesSQL = "SELECT COUNT(ID) FROM Usr_Approvals WHERE WhoApprovesMe = '{0}' OR WAM2 = '{0}' OR WAM3 = '{0}' OR WAM4 = '{0}'".format(userToCheck)
  
  ## New verison using 'Usr_HODapprovals' table
  ##tmpCountAppearancesSQL = "SELECT ISNULL(STRING_AGG(EMailExternal, '; '), 'Matt.Pattison@thackraywilliams.com')  FROM Users WHERE Code IN (SELECT UserCode FROM Usr_HODapprovals WHERE FeeEarnerCode = '{FeeEarnerRef}')".format(FeeEarnerRef = userToCheck)
  #tmpCountAppearancesSQL = "SELECT COUNT(UserCode) FROM Usr_HODapprovals WHERE UserCode = '{HODref}'".format(HODref=userToCheck)
  # Just realising an issue with the above is that we're stating that if user doesn't have anyone setup for them to approve (in said new table), then they're not an approver
  # which isn't necessarily the case... Instead it may serve us better to work from the 'Locks'/'Keys' table, as approval users will have access to the HOD screen via having Key to the Lock/screen
  HODusersSQL = "SELECT STRING_AGG(UserRef, ' | ') FROM Keys WHERE LockRef = ( SELECT Code FROM Locks WHERE Description = 'XAML_Screen_HOD_AccessOnly')"
  HODusers = runSQL(HODusersSQL)

  if userToCheck in HODusers:
    return True
  else:
    return False
  

def canApproveSelf(userToCheck):
  # This function will return boolean (True or False) to indicate whether the passed user can approve themselves (by checking if users email address is in the appover list)
  
  # get email address of user
  userEmail = runSQL("SELECT EMailExternal FROM Users WHERE Code = '{0}'".format(userToCheck))
  tmpHODemails = getUsersApproversEmail(forUser = userToCheck)
  
  if userEmail in tmpHODemails:
    return True
  else:
    return False


def canUserApproveFeeEarner(UserToCheck, FeeEarner):
  # This function will return boolean (True or False) to indicate whether the passed 'UserToCheck' can Approve the passed 'FeeEarner'
  
  # get email address of user
  userEmail = runSQL("SELECT EMailExternal FROM Users WHERE Code = '{0}'".format(UserToCheck))
  tmpHODemails = getUsersApproversEmail(forUser = FeeEarner)
  
  if userEmail in tmpHODemails:
    return True
  else:
    return False

def canUserReviewFiles(userToCheck):
  # This function will return a bool value for the ability of the fee earner to conduct file reviews
  ## GRADES             | CAN DO FR?
  ## Associate          | Yes
  ## Equity Partner     | Yes
  ## Partner            | Yes
  ## Head of Department | Yes
  ## Senior Solicitor   | Yes
  ## Solicitor          | Yes
  ## Trainee            | No
  ## ILEX               | No
  ## PA                 | No
  ## Secretary          | No
  ## Management Team    | No
  ## Other              | No

  # MP: added override for IT and Risk to be able to do:
  if userToCheck in RiskAndITUsers:
    return True

  gradesThatCanFR = ['Associate', 'Equity Partner', 'Partner', 'Head of Department', 'Senior Solicitor', 'Solicitor']
  gradeCheck = runSQL("""SELECT G.Description AS Grade
                         FROM Users U
                         LEFT OUTER JOIN Grades G ON U.UserGrade = G.Code
                         LEFT OUTER JOIN Departments D ON U.Department = D.Code
                         WHERE U.Code = '{0}'
                         ORDER BY D.Description, U.FullName""".format(userToCheck))
  
  # since 22nd July 2025, Amy wants restrictions lifted, and left to individual departments to manage
  # Therefore, as we still don't want trainees or lower doing, we still need to check users grade against our list of 'approved' grades:
  return True if gradeCheck in gradesThatCanFR else False

  # Original Rules - According to Ali - only 'Partner', 'Equity Partner' can do, and if Private Client, also 'Associate'
  #departmentCheck = runSQL("""SELECT D.Description AS Department
  #                            FROM Users U
  #                            LEFT OUTER JOIN Departments D ON U.Department = D.Code
  #                            WHERE U.Code = '{0}'
  #                            ORDER BY D.Description, U.FullName""".format(userToCheck))
    
  # Check if the department is not "Private Client" 
  # and gradeCheck matches "associate", "partner", or "equity partner"
  #if departmentCheck == "Private Client" and gradeCheck in ["Associate", "Partner", "Equity Partner"]:
  #  return True
  #elif departmentCheck != "Private Client" and gradeCheck in ["Partner", "Equity Partner"]:
  #  return True
  #else:
  #  return False



def getUsersApproversEmail(forUser):
  # This function will return a list of email addresses of the passed forUser
  
  # old version using 'Usr_Approvals' table
  #hodEmailSQL = """SELECT STRING_AGG(EMailExternal, '; ') FROM Users WHERE Code IN (
  #                SELECT 'Who' = WhoApprovesMe FROM Usr_Approvals WHERE UserCode = '{0}' 
  #                UNION SELECT 'Who' = WAM2 FROM Usr_Approvals WHERE UserCode = '{0}' 
  #                UNION SELECT 'Who' = WAM3 FROM Usr_Approvals WHERE UserCode = '{0}' 
  #                UNION SELECT 'Who' = WAM4 FROM Usr_Approvals WHERE UserCode = '{0}')""".format(forUser)

  # New verison running off 'Usr_HODapprovals' table
  hodEmailSQL = "SELECT ISNULL(STRING_AGG(EMailExternal, '; '), 'Matt.Pattison@thackraywilliams.com')  FROM Users WHERE Code IN (SELECT UserCode FROM Usr_HODapprovals WHERE FeeEarnerCode = '{FeeEarnerRef}')".format(FeeEarnerRef = forUser)
  hodEmail = runSQL(hodEmailSQL)
  #hodEmail = runSQL(hodEmailSQL, True, 'There was an error getting approval users email address...', 'DEBUGGING - getUsersApproversEmail')
  return hodEmail


def HOD_Approves_Item(my_mraID, myEntRef, myMatNo, myMRADesc):
  # This is a generic function to 'approve' an item where we pass in the parameters (better for re-use, instead of copying and pasting)
  # This assumes current user is HOD/Approver - addresses email to Matter Fee Earner, and copies in 'current user'
  errorCount = 0
  errorMessage = ""
  
  # get / form input variables (at global header level, so accessible anywhere)
  tmpOurRef = tb_OurRef.Text
  tmpToUserName = tb_FE_Forename.Text
  tmpMatDesc = tb_MatterDesc.Text
  tmpClName = tb_ClientName.Text
  tmpEmailTo = tb_FE_Email.Text
  tmpEmailCC = get_user_email(_tikitUser)
  tmpAddtl1 = myMRADesc.replace("'", "''")
  tmpAddtl2 = "High"

  # generate SQL to approve
  approveSQL = """UPDATE Usr_MRAv2_MatterHeader SET ApprovedByHOD = 'Y' 
                  WHERE mraID = {mraID} AND EntityRef = '{entRef}' AND MatterNo = {matNo}""".format(mraID=my_mraID, entRef=myEntRef, matNo=myMatNo)
  try:
    _tikitResolver.Resolve("[SQL: {0}]".format(approveSQL))
  except:
    errorCount -= 1
    errorMessage += " - couldn't mark the selected item as approved\n" + str(approveSQL)
  
  # get SQL to Unlock matter
  #unlockCode = "EXEC TW_LockHandler '" + myEntRef + "', " + str(myMatNo) + ", 'LockedByRiskDept', 'UnLock'"
  unlockCode = "EXEC TW_LockHandler '{entRef}', {matNo}, 'LockedByRiskDept', 'UnLock'".format(entRef=myEntRef, matNo=myMatNo)
  # run unlock code
  runSQL(codeToRun=unlockCode, showError=True, 
         errorMsgText="There was an error unlocking the matter, after approval. Please check the matter is unlocked and if not, unlock manually using the following SQL:\n{0}".format(unlockCode), errorMsgTitle="Error: Unlocking Matter after Approval...")
 
  # now insert a record into MRA Events table to trigger email to FE
  if insert_into_MRAEvents(userRef=_tikitUser, triggerText='HOD_Approved_MRA', ov_ID=my_mraID, 
                           emailTo=tmpEmailTo, emailCC=tmpEmailCC, toUserName=tmpToUserName, 
                           ourRef=tmpOurRef, matterDesc=tmpMatDesc, clientName=tmpClName, 
                           addtl1=tmpAddtl1, addtl2=tmpAddtl2) == False:
    errorCount -= 1
    errorMessage += " - couldn't send the 'HOD Approved' Task Centre confirmation email to FE\n"

  if errorCount < 0:
    MessageBox.Show("The following error(s) were encountered:\n" + errorMessage + "\n\nPlease screenshot this message and send to IT.Support@thackraywilliams.com to investigate", "Error: Approve High-Risk Matter...")
    return errorCount
  else:
    createNewMRA_BasedOnCurrent(idItemToCopy=my_mraID, entRef=myEntRef, matNo=myMatNo)
    MessageBox.Show("Successfully Approved the Matter Risk Assessment (MRA) and Unlocked the matter.\n\nA copy of the MRA has been made, to be completed by the Fee Earner within 4 weeks", "Approve High-Risk Matter...")
    return 1
  return 0


def createNewMRA_BasedOnCurrent(idItemToCopy, entRef, matNo):
  # this function will duplicate the MRA specified by 'idItemToCopy' (which is the mraID of the MRA to copy),
  # and link the new MRA to the same matter (using entRef and matNo) and copy across the selected Answers used in the original MRA
  # - set status to 'Draft' and score to 0, and expiry date to 4 weeks from today.
  #! Returns mraID of newly added/duplicated row, or -1 if error creating new MRA header row, or -2 if error copying details across

  # generate new Name (concatenating Template Name with next number in brackets, based on count of existing MRAs with same Template for the same matter +1):
  newNameSQL = """SELECT CONCAT(TD.Name, ' (', 
                            CONVERT(nvarchar, (
                                SELECT ISNULL(COUNT(MH.TemplateID), 0) + 1 FROM Usr_MRAv2_MatterHeader MH 
                                WHERE MH.EntityRef = '{entRef}' AND MH.MatterNo = {matNo} AND MH.TemplateID = TD.TemplateID)),
                          ')')
                  FROM Usr_MRAv2_TemplateDetails TD 
                  WHERE TD.TemplateID = (SELECT TemplateID FROM Usr_MRAv2_MatterHeader WHERE mraID = {mraID})""".format(entRef=entRef, matNo=matNo, mraID=idItemToCopy)  

  newName = runSQL(newNameSQL)

  # create a duplicate row in the 'Matter Header' table (re-setting score and status and expiry date etc, but copying across the TemplateID and linking to the same matter)
  newHeaderRowSQL = """INSERT INTO Usr_MRAv2_MatterHeader (EntityRef, MatterNo, TemplateID, Name, RiskRating, ApprovedByHOD, ExpiryDate, DateAdded, mraID, Status) 
                       OUTPUT INSERTED.mraID 
                       SELECT '{entRef}', {matNo}, TemplateID, '{passedName}', RiskRating, 'N', DATEADD(WEEK, 4, ExpiryDate),
                       GETDATE(), (SELECT MAX(mraID) + 1 FROM Usr_MRAv2_MatterHeader), 'Draft' FROM Usr_MRAv2_MatterHeader WHERE mraID = {idItemToCopy}  
                    """.format(entRef=entRef, matNo=matNo, passedName=newName, idItemToCopy=idItemToCopy)

  try:
    newMRAID = _tikitResolver.Resolve("[SQL: {0}]".format(newHeaderRowSQL))
  except:
    MessageBox.Show("There was an error creating a new MRA, using SQL:\n" + str(newHeaderRowSQL), "Error: Duplicate selected item...")
    return -1

  # finally, copy across the Questions and Answers from the previous MRA (except for any free text answers, which we will reset to blank) - we can link them to the new MRA using the returned mraID from above, which is why we needed to use 'OUTPUT' in the SQL above to return the new mraID
  copyDetailSQL = """INSERT INTO Usr_MRAv2_MatterDetails (EntityRef, MatterNo, mraID, QuestionID, AnswerID, Score, Comments, DisplayOrder, QuestionGroup, EmailComment) 
                     SELECT EntityRef, MatterNo, {NEWmraID}, QuestionID, AnswerID, Score, Comments, DisplayOrder, QuestionGroup, EmailComment
                     FROM Usr_MRAv2_MatterDetails WHERE mraID = {idToCopy}""".format(NEWmraID=newMRAID, idToCopy=idItemToCopy)

  try:
    _tikitResolver.Resolve("[SQL: {0}]".format(copyDetailSQL))
  except:
    MessageBox.Show("There was an error copying the MRA details, using SQL:\n" + str(copyDetailSQL), "Error: Duplicate selected item...")
    return -2
  return newMRAID


def createNewFR_BasedOnCurrent(idItemToCopy, entRef, matNo):
  # this function will duplicate the File Review specified by 'idItemToCopy' (which is the ID of the File Review to copy),
  # and link the new File Review to the same matter (using entRef and matNo) and copy across the selected Answers used in the original File Review
  # - set status to 'Draft' and expiry date to 4 weeks from today.

  # generate new Name (concatenating Template Name with next number in brackets, based on count of existing MRAs with same Template for the same matter +1):
  newNameSQL = """SELECT CONCAT(TT.TypeName, ' (', 
                            CONVERT(nvarchar, (
                                SELECT ISNULL(COUNT(MRAO.TypeID), 0) + 1 FROM Usr_MRA_Overview MRAO 
                                WHERE MRAO.EntityRef = '{entRef}' AND MRAO.MatterNo = {matNo} AND MRAO.TypeID = TT.TypeID)),
                          ')')
                  FROM Usr_MRA_TemplateTypes TT
                  WHERE TT.TypeID = (SELECT TypeID FROM Usr_MRA_Overview WHERE ID = {mraID})""".format(entRef=entRef, matNo=matNo, mraID=idItemToCopy)  

  newName = runSQL(newNameSQL)

  # create a duplicate row in the 'Matter Header' table (re-setting score and status and expiry date etc, but copying across the TemplateID and linking to the same matter)
  newHeaderRowSQL = """INSERT INTO Usr_MRA_Overview (EntityRef, MatterNo, TypeID, LocalName, RiskRating, ApprovedByHOD, ExpiryDate, DateAdded, Status) 
                       OUTPUT INSERTED.ID 
                       SELECT '{entRef}', {matNo}, TypeID, '{passedName}', RiskRating, 'N', DATEADD(WEEK, 4, ExpiryDate),
                       GETDATE(), 'Draft' FROM Usr_MRA_Overview WHERE ID = {idItemToCopy} 
                    """.format(entRef=entRef, matNo=matNo, passedName=newName, idItemToCopy=idItemToCopy)

  try:
    newID = _tikitResolver.Resolve("[SQL: {0}]".format(newHeaderRowSQL))
  except:
    MessageBox.Show("There was an error creating a new MRA, using SQL:\n" + str(newHeaderRowSQL), "Error: Duplicate selected item...")
    return -1

  # finally, copy across the Questions and Answers from the previous MRA (except for any free text answers, which we will reset to blank) - we can link them to the new MRA using the returned mraID from above, which is why we needed to use 'OUTPUT' in the SQL above to return the new mraID
  copyDetailSQL = """INSERT INTO Usr_MRA_Detail (EntityRef, MatterNo, OV_ID, QuestionID, AnswerListToUse, SelectedAnswerID, CurrentAnswerScore, Notes, DisplayOrder, QGroupID, EmailComment) 
                     SELECT EntityRef, MatterNo, {newID}, QuestionID, AnswerListToUse, SelectedAnswerID, CurrentAnswerScore, Notes, DisplayOrder, QGroupID, EmailComment
                     FROM Usr_MRA_Detail WHERE OV_ID = {idToCopy}""".format(newID=newID, idToCopy=idItemToCopy)

  try:
    _tikitResolver.Resolve("[SQL: {0}]".format(copyDetailSQL))
  except:
    MessageBox.Show("There was an error copying the MRA details, using SQL:\n" + str(copyDetailSQL), "Error: Duplicate selected item...")
    return -2
  return newID

# # # # #   END OF:   G E N E R I C   F U N C T I O N S   # # # # #

def dg_MRAFR_ReSaveEmailToCase(s, event):
  # This function will save a copy of the selected Matter Risk Assessment or File Review email into the case.
  # There will need to be 5 Task Centre tasks, as content of email is different for each.
  # 2 for Matter Risk Assessments ('Standard'/Med; and 'High Risk' - though need to see if HOD approved*); and 
  # 3 for File Reviews ('No Corrective Actions', 'Corrective Actions - with FE', 'Corrective Actions - Completed')
  # *note: there is a separate email for if the HOD has approved High Risk matter, so need to check and do applicable email
  # *eg: if HOD not yet approved, save the 'request for approval' email, else (if IS approved already), save the 'HOD Approved' version
  # * I do wonder if users actually want to trigger emails too / or want option of which to re-save in above example...

  MessageBox.Show("This function is coming soon...", "Resave selected item to case - Coming soon...")
  return

  if dg_MRAFR.SelectedIndex == -1:
    MessageBox.Show("Nothing selected to save to case!", "Error: Resave selected item to case...")
    return

  tmpID = dg_MRAFR.SelectedItem['mraID']
  tmpType = dg_MRAFR.SelectedItem['Type']
  tmpName = dg_MRAFR.SelectedItem['Name']
  tmpStatus = dg_MRAFR.SelectedItem['Status']

  # alert user and exit funtion if status is 'Draft' or 'In Progress'
  if 'Draft' in tmpStatus or 'In Progress' in tmpStatus:
    MessageBox.Show("You cannot save an item that has a Status of 'Draft' or 'In Progress'!", "Error: Resave selected item to case...")
    return

  # we need to determine which task to run based on the type
  if 'File Review' in tmpType:
    # is a FR...
    # set the appropriate triggerText according to Status (options are: 'FR_CorrectiveActions_Complete'; 'FR_CorrectiveActions_WithFE'; 'FR_Complete')
    # ! might need to check for whether there are any corrective actions to complete, so we can set the correct triggerText
    
    taskTriggerText = ""
  else:
    # is a MRA...
    # set the appropriate triggerText according to Status (options are: 'Submit_MRA', 'Submit_MRA_HighRisk', 'HOD_Approved_MRA')
    taskTriggerText = ""


  return






]]>
    </Init>
    <Loaded>
      <![CDATA[
ti_Main = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_Main')
ti_MRA = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_MRA')
ti_FR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_FR')

tb_FERef = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_FERef')
tb_FE_Forename = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_FE_Forename')
tb_FE_Email = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_FE_Email')
#lbl_HOD_Email = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_HOD_Email')
tb_MatterDesc = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MatterDesc')
tb_CaseTypeRef = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_CaseTypeRef')
tb_CaseType = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_CaseType')
tb_OurRef = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_OurRef')
tb_ClientName = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_ClientName')
tb_CurrentUserName = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_CurrentUserName')
tb_CurrentUserCode = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_CurrentUserCode')

## O V E R V I E W   - TAB ##
dg_MRAFR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_MRAFR')
dg_MRAFR.SelectionChanged += dg_MRAFR_SelectionChanged
lbl_MRAFR_ID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRAFR_ID')
lbl_MRAFR_Name = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRAFR_Name')

# Note: 'New' button is now wired as a 'template', so 
btnNew = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btnNew')
btnNew.Click += btnNew_Click
popTemplates = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'popTemplates')
popTemplates.Closed += popTemplates_Closed
icTemplates = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'icTemplates')
btnCancel = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btnCancel')
btnCancel.Click += CancelPopup_Click

btn_CopySelected_MRAFR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_CopySelected_MRAFR')
btn_CopySelected_MRAFR.Click += btn_CopySelected_MRAFR_Click
btn_View_MRAFR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_View_MRAFR')
btn_View_MRAFR.Click += dg_MRAFR_ViewSelected
btn_Edit_MRAFR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Edit_MRAFR')
btn_Edit_MRAFR.Click += btn_Edit_MRAFR_Click
btn_HOD_Approval_MRA = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_HOD_Approval_MRA')
btn_HOD_Approval_MRA.Click += btn_HOD_Approval_MRA_Click
btn_DeleteSelected_MRAFR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_DeleteSelected_MRAFR')
btn_DeleteSelected_MRAFR.Click += dg_MRAFR_DeleteSelected
#btn_DeleteSelected_MRAFR.Click += lambda s, event: (test_dgItem_DeleteSelected(childDeletes = [{'table': 'Usr_MRA_Detail','foreignKey': 'OV_ID','conditions': {'EntityRef': _tikitEntity,'MatterNo': _tikitMatter}}], dgControl=dg_MRAFR, tableToUpdate='Usr_MRA_Overview', primaryKeyColName='ID', displayOrderColName='', keyColumns={'EntityRef': _tikitEntity, 'MatterNo': _tikitMatter}, selectedColumns={'ID': 'ID', 'Name': 'Desc'}, typeFilterColName='', typeFilterVal=0), dg_MRAFR_Refresh(s, event))
# Louis version - but can't remember why this isn't 'active'
btn_RegenerateEmailForFile = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_RegenerateEmailForFile')
btn_RegenerateEmailForFile.Click += dg_MRAFR_ReSaveEmailToCase

lbl_OV_RiskStatus = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_OV_RiskStatus')
lbl_RiskScore_AdvisoryText = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_RiskScore_AdvisoryText')
tb_NoMRAFR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_NoMRAFR')
sep_Delete = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'sep_Delete')

## OVERVIEW - CORRECTIVE ACTIONS ##
dgCA_Overview = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dgCA_Overview')
dgCA_Overview.SelectionChanged += dgCA_Overview_SelectionChanged
dgCA_Overview.CellEditEnding += dgCA_Overview_CellEditEnding
tb_NoCAs = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_NoCAs')
btn_Mark_CA_Complete = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Mark_CA_Complete')
btn_Mark_CA_Complete.Click += dgCA_Overview_ToggleComplete
btn_View_CA_onFR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_View_CA_onFR')
btn_View_CA_onFR.Click += dg_CA_Overview_ViewOnFileReview
tb_Current_CA = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_Current_CA')
tb_Current_CA_Complete = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_Current_CA_Complete')
tb_CurrUser = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_CurrUser')
tb_CurrUserName = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_CurrUserName')
opt_CA_ViewNotComplete = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'opt_CA_ViewNotComplete')
opt_CA_ViewNotComplete.Click += dgCA_Overview_Refresh
opt_CA_ViewComplete = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'opt_CA_ViewComplete')
opt_CA_ViewComplete.Click += dgCA_Overview_Refresh
opt_CA_ViewAll = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'opt_CA_ViewAll')
opt_CA_ViewAll.Click += dgCA_Overview_Refresh
tb_CATaken = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_CATaken')
btn_UpdateReviewerWithActionTaken = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_UpdateReviewerWithActionTaken')
btn_UpdateReviewerWithActionTaken.Click += FR_UpdateReviewerWithActionTaken_Click


##   M A T T E R   R I S K   A S S E S S M E N T   - TAB ##
lbl_MRA_Status = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRA_Status')
stk_RiskInfo = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'stk_RiskInfo')
lbl_MRA_Score = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRA_Score')
lbl_MRA_RClabel = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRA_RClabel')
lbl_MRA_RiskCategory = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRA_RiskCategory')
lbl_MRA_RiskCategoryID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRA_RiskCategoryID')
lbl_TotalQs = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_TotalQs')
lbl_TotalAnswered = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_TotalAnswered')

btn_MRA_SaveAsDraft = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_MRA_SaveAsDraft')
btn_MRA_SaveAsDraft.Click += btn_MRA_SaveAsDraft_Click
btn_MRA_Submit = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_MRA_Submit')
btn_MRA_Submit.Click += btn_MRA_Submit_Click
btn_MRA_BackToOverview = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_MRA_BackToOverview')
btn_MRA_BackToOverview.Click += btn_MRA_BackToOverview_Click
btn_HOD_Approval_MRA1 = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_HOD_Approval_MRA1')
btn_HOD_Approval_MRA1.Click += btn_HOD_Approval_MRA1_Click

tb_DaysUntilLocked = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_DaysUntilLocked')
tb_MatterWillBeLockedMsg = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MatterWillBeLockedMsg')

lbl_TimeEntered = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_TimeEntered')

lbl_MRA_ID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRA_ID')
lbl_MRA_Name = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRA_Name')
lbl_MRA_TemplateID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRA_TemplateID')
lbl_ScoreTrigger_Medium = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_ScoreTrigger_Medium')
lbl_ScoreTrigger_High = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_ScoreTrigger_High')

tb_NoMRA_Qs = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_NoMRA_Qs')
grid_MRA = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'grid_MRA')
dg_MRA = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_MRA')
dg_MRA.SelectionChanged += dg_MRA_SelectionChanged

grp_MRA_SelectedQuestionArea = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'grp_MRA_SelectedQuestionArea')
chk_MRA_AutoSelectNext = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'chk_MRA_AutoSelectNext')
cbo_MRA_SelectedComboAnswer = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cbo_MRA_SelectedComboAnswer')
#cbo_MRA_SelectedComboAnswer.DropDownClosed += MRA_SaveAnswer
cbo_MRA_SelectedComboAnswer.SelectionChanged += cbo_MRA_SelectedComboAnswer_SelectionChanged
tb_MRA_QNotes = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MRA_QNotes')
#tb_MRA_QNotes.LostFocus += MRA_SaveAnswer
btn_MRA_SaveAnswer = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_MRA_SaveAnswer')
btn_MRA_SaveAnswer.Click += btn_MRASaveAnswer_Click


##   F I L E   R E V I E W   - TAB ##
txt_FR_Status = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'txt_FR_Status')
tb_TotalQs_FR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_TotalQs_FR')
tb_TotalAnswered_FR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_TotalAnswered_FR')
tb_TotalCAs_FR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_TotalCAs_FR')
tb_TotalOSCAs_FR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_TotalOSCAs_FR')
btn_FR_BackToOverview = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_FR_BackToOverview')
btn_FR_BackToOverview.Click += FR_BackToOverview
btn_FR_Submit = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_FR_Submit')
btn_FR_Submit.Click += btn_FR_Submit_Click
lbl_FR_ID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_FR_ID')
lbl_FR_Name = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_FR_Name')
tb_NoFR_Qs = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_NoFR_Qs')
dg_FR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_FR')
#dg_FR.CellEditEnding += FR_CellEditEnding
dg_FR.SelectionChanged += FR_SelectionChanged
lbl_FR_DGID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_FR_DGID')
lbl_FR_CurrVal = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_FR_CurrVal')
chk_FR_AutoSelectNext = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'chk_FR_AutoSelectNext')
lbl_TimeEnteredFR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_TimeEnteredFR')

opt_EditModeFR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'opt_EditModeFR')
opt_ViewModeFR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'opt_ViewModeFR')

opt_Yes = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'opt_Yes')
opt_Yes.Click += FR_optYes_Clicked
opt_No = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'opt_No')
opt_No.Click += FR_optNo_Clicked
opt_NA = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'opt_NA')
opt_NA.Click += FR_optNA_Clicked

chk_FR_AllowsNA = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'chk_FR_AllowsNA')
lbl_FR_CAtrigger = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_FR_CAtrigger')
chk_FR_AllowsNotes = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'chk_FR_AllowsNotes')
txt_FR_QComment = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'txt_FR_QComment')
#txt_FR_QComment.TextChanged += FR_QComment_Save
txt_FR_QComment.LostFocus += FR_QComment_Save

# testing if we can action cell update based off direct combo box access
#cbo_FR_DGAnswer = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cbo_FR_DGAnswer')
#cbo_FR_DGAnswer.SelectionChanged += FR_DG_AnswerChanged
#cbo_FR_DGAnswer.SelectedValueChanged += FR_DG_AnswerChanged
# 'SelectionChanged' is correct for a Combo box, but we still get error that 'none type' doesn't have this event
# This is a shame, it appears we cannot get handle on item at run time, nor can we explicitly use event triggers on XAML (these crash P4W)
# Therefore, really need to think about how the 'auto-advance' triggers



## C O R R E C T I V E    A C T I O N S   ##
#grd_Main = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'grd_Main')
#grd_Col0 = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'grd_Col0')
stk_CorrectiveActions = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'stk_CorrectiveActions')

lbl_CorrActionID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_CorrActionID')
lbl_CorrAct_QText = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_CorrAct_QText')
txt_CorractiveActionNeeded = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'txt_CorractiveActionNeeded')
txt_CorractiveActionNeeded.LostFocus += txt_CorractiveActionNeeded_LostFocus
txt_CorractiveActionTaken = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'txt_CorractiveActionTaken')
txt_CorractiveActionTaken.LostFocus += txt_CorractiveActionTaken_LostFocus
chk_CorrectiveActionPassed = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'chk_CorrectiveActionPassed')
chk_CorrectiveActionPassed.Click += chk_CorrectiveActionPassed_Click
btn_CorrAction_Save = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_CorrAction_Save')
btn_CorrAction_Save.Click += btn_CorrectiveAction_Save_Clicked


## C A S E   D O C S    C O N T R O L S   ##
opt_CaseDocs_Entity = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'opt_CaseDocs_Entity')
opt_CaseDocs_Entity.Click += opt_EntityOrMatterDocs_Clicked
opt_CaseDocs_Matter = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'opt_CaseDocs_Matter')
opt_CaseDocs_Matter.Click += opt_EntityOrMatterDocs_Clicked
dg_CaseManagerDocs = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_CaseManagerDocs')
dg_CaseManagerDocs.SelectionChanged += CaseDoc_SelectionChanged
btn_OpenCaseDoc = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_OpenCaseDoc')
btn_OpenCaseDoc.Click += open_Selected_CaseDoc
cbo_AgendaName = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cbo_AgendaName')
cbo_AgendaName.SelectionChanged += refresh_CaseDocs

#tog_TEST_CorrActions = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tog_TEST_CorrActions')
#tog_TEST_CorrActions.Click += tog_Test_CorrectiveActions

# call 'On Load' event
myOnLoadEvent(_tikitSender, 'onLoad')

]]>
    </Loaded>
  </RiskMatterV2>
</tfb>
