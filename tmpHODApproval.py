<?xml version="1.0" encoding="UTF-8"?>
<tfb>
  <RiskPractice>
    <Init>
      <![CDATA[
import clr

clr.AddReference('mscorlib')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('System.Windows.Forms')

from System import DateTime
from System.Diagnostics import Process
from System.Globalization import DateTimeStyles
from System.Collections.Generic import Dictionary
from System.Windows import Controls, Forms, LogicalTreeHelper
from System.Windows import Data, UIElement, Visibility, Window
from System.Windows.Controls import Button, Canvas, GridView, GridViewColumn, ListView, Orientation
from System.Windows.Data import Binding, CollectionView, ListCollectionView, PropertyGroupDescription
from System.Windows.Forms import SelectionMode, MessageBox, MessageBoxButtons, DialogResult
from System.Windows.Input import KeyEventHandler
from System.Windows.Media import Brush, Brushes

## GLOBAL VARIABLES ##
preview_MRA = []    # To temp store table for previewing Matter Risk Assessment
previewFR = []      # To temp store table for previewing File Review


# # # #   O N   L O A D   E V E N T   # # # #
def myOnLoadEvent(s, event):
  
  populateUsersList(s, event)
  setCurrentUser(s, event)
  
  whoCanChangeCurrentUser = ['LD1', 'MP', 'NB']
  if _tikitUser in whoCanChangeCurrentUser:
    stk_CurrentUser.Visibility = Visibility.Visible
  else:
    stk_CurrentUser.Visibility = Visibility.Collapsed

  # populate drop-downs
  populate_FeeEarnersList(s, event)
  refresh_Preview_MRA(s, event)
  
  #MessageBox.Show("Hello world! - OnLoad Finished")
  return



class UsersList(object):
  def __init__(self, myFEName, myBranch, myOurRef, myMRAname, myMRAexp, myOV_ID, myEntRef, myMatNo, myType, myClName, myMatDesc):
    self.feName = myFEName
    self.feBranch = myBranch
    self.ourRef = myOurRef
    self.mraName = myMRAname
    self.mraExpiryDate = myMRAexp
    self.ov_ID = myOV_ID
    self.entRef = myEntRef
    self.matNo = myMatNo
    self.AType = myType
    self.clName = myClName
    self.matDesc = myMatDesc
    return
    
  def __getitem__(self, index):
    if index == 'FEName':
      return self.feName
    elif index == 'Branch':
      return self.feBranch
    elif index == 'OurRef':
      return self.ourRef
    elif index == 'MRAName':
      return self.mraName
    elif index == 'MRA Expiry':
      return self.mraExpiryDate
    elif index == 'OV_ID':
      return self.ov_ID
    elif index == 'EntityRef':
      return self.entRef
    elif index == 'MatterNo':
      return self.matNo
    elif index == 'Type':
      return self.AType
    elif index == 'ClientName':
      return self.clName
    elif index == 'MatDesc':
      return self.matDesc


def populate_FeeEarnersList(s, event): 

  # FOLLOWING CODE IS DIRECTLY FROM WIP SCREEN
  currentUser = cbo_User.SelectedItem['Code']
  #user_Dept = runSQL("SELECT Department FROM Users WHERE Code = '" + currentUser + "'", False, "", "")
  userIsHOD = isUserAnApprovalUser(userToCheck = currentUser)
  myFEitems = []
  
  ## if Active user is a HOD and a Team lead, then include their name first in the list
  #if userIsHOD == False:
  #  tb_NoFEs.Visibility = Visibility.Visible
  #  stk_SelectedMRAHeader_NoData.Visibility = Visibility.Visible
  #  stk_SelectedMRAHeader.Visibility = Visibility.Collapsed
  #  dg_MRAReview.Visibility = Visibility.Collapsed
  #  dg_FeeEarners.Visibility = Visibility.Collapsed
  #  return
  #else:
  #  tb_NoFEs.Visibility = Visibility.Collapsed
  #  stk_SelectedMRAHeader_NoData.Visibility = Visibility.Collapsed
  #  stk_SelectedMRAHeader.Visibility = Visibility.Visible
  #  dg_MRAReview.Visibility = Visibility.Visible
  #  dg_FeeEarners.Visibility = Visibility.Visible

  mySQL = """SELECT '0-FeeEarner' = U.FullName, 
                    '1-Branch' = B.Description, 
                    '2-OurRef' = LEFT(MRAO.EntityRef, 3) + RIGHT(MRAO.EntityRef, 4) + '/' + CONVERT(nvarchar, MRAO.MatterNo, 1), 
                    '3-MRA Name' = MRAO.LocalName, 
                    '4-Expiry Date' = MRAO.ExpiryDate, 
                    '5-OV_ID' = MRAO.ID, 
                    '6-EntityRef' = MRAO.EntityRef, 
                    '7-MatterNo' = MRAO.MatterNo, 
                    '8-CaseType' = CT.Description, 
                    '9-Type' = CASE WHEN '{0}' IN (SELECT HA.UserCode FROM Usr_HODapprovals HA WHERE HA.FeeEarnerCode = U.Code AND HA.Type = 'Main') THEN '01) Main' 
                                    WHEN '{0}' IN (SELECT HA.UserCode FROM Usr_HODapprovals HA WHERE HA.FeeEarnerCode = U.Code AND HA.Type = 'Holiday Cover') THEN '02) Holiday Cover' ELSE '' END, 
                    '10-MatterDesc' = M.Description, 
                    '11-EntName' = E.LegalName 
        FROM Usr_MRA_Overview MRAO  
          INNER JOIN Matters M ON MRAO.EntityRef = M.EntityRef AND MRAO.MatterNo = M.Number 
          LEFT OUTER JOIN Users U ON M.FeeEarnerRef = U.Code 
          LEFT OUTER JOIN Branches B ON U.UsualOffice = B.Code
          INNER JOIN CaseTypes CT ON M.CaseTypeRef = CT.Code
          INNER JOIN CaseTypeGroups CTG ON CT.CaseTypeGroupRef = CTG.ID
          LEFT OUTER JOIN Usr_Approvals App ON U.Code = App.UserCode
          LEFT OUTER JOIN Entities E ON MRAO.EntityRef = E.Code
        WHERE MRAO.ApprovedByHOD = 'N' AND MRAO.RiskRating = 3 AND MRAO.Status = 'Complete'
          AND U.Code IN (SELECT FeeEarnerCode FROM Usr_HODapprovals WHERE UserCode = '{0}')
        ORDER BY '9-Type', '0-FeeEarner', '4-Expiry Date'
        """.format(currentUser)

  _tikitDbAccess.Open(mySQL)

  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          iFE = '-' if dr.IsDBNull(0) else dr.GetString(0)
          iBranch = '-' if dr.IsDBNull(1) else dr.GetString(1)
          iOurRef = '' if dr.IsDBNull(2) else dr.GetString(2)
          iMRAName = 0 if dr.IsDBNull(3) else dr.GetString(3)
          iMRAExpiry = 0 if dr.IsDBNull(4) else dr.GetValue(4)
          iOVID = 0 if dr.IsDBNull(5) else dr.GetValue(5)
          iEntRef = 0 if dr.IsDBNull(6) else dr.GetString(6)
          iMatNo = 0 if dr.IsDBNull(7) else dr.GetValue(7)
          iType = '' if dr.IsDBNull(9) else dr.GetString(9)
          iMatDesc = '' if dr.IsDBNull(10) else dr.GetString(10)
          iClName = '' if dr.IsDBNull(11) else dr.GetString(11)
          
          myFEitems.append(UsersList(iFE, iBranch, iOurRef, iMRAName, iMRAExpiry, iOVID, iEntRef, iMatNo, iType, iClName, iMatDesc))
      
    dr.Close()
  _tikitDbAccess.Close()
  
  # now group Fee Earners by Branch and set THAT object as the ItemsSource...
  tmpC = ListCollectionView(myFEitems)
  #tmpC.GroupDescriptions.Add(PropertyGroupDescription("feBranch"))
  tmpC.GroupDescriptions.Add(PropertyGroupDescription("AType"))
  tmpC.GroupDescriptions.Add(PropertyGroupDescription("feName"))
  
  dg_FeeEarners.ItemsSource = tmpC
  
  if dg_FeeEarners.Items.Count == 0:
    tb_NoFEs.Visibility = Visibility.Visible
    stk_SelectedMRAHeader_NoData.Visibility = Visibility.Visible
    stk_SelectedMRAHeader.Visibility = Visibility.Collapsed
    dg_MRAReview.Visibility = Visibility.Collapsed
    dg_FeeEarners.Visibility = Visibility.Collapsed
  else:
    tb_NoFEs.Visibility = Visibility.Collapsed
    stk_SelectedMRAHeader_NoData.Visibility = Visibility.Collapsed
    stk_SelectedMRAHeader.Visibility = Visibility.Visible
    dg_MRAReview.Visibility = Visibility.Visible
    dg_FeeEarners.Visibility = Visibility.Visible
  return
  
  
  

# # # #   P R E V I E W   M A T T E R   R I S K   A S S E S S M E N T   TAB # # # #

class review_MRA(object):
  def __init__(self, myOrder, myQuestion, myAnswerText, myFEnotes, myEC, myGroup):
    self.DOrder = myOrder
    self.QuestionText = myQuestion
    self.LUP_AnswerText = myAnswerText
    self.FENotes = myFEnotes
    self.LUP_EmailComment = myEC
    self.QGroup = myGroup
    return
    
  def __getitem__(self, index):
    if index == 'Order':
      return self.DOrder
    elif index == 'Question':
      return self.QuestionText
    elif index == 'AnswerText':
      return self.LUP_AnswerText
    elif index == 'FEnotes':
      return self.FENotes
    elif index == 'EmailComment':
      return self.LUP_EmailComment
    elif index == 'QGroup':
      return self.QGroup
    else:
      return ''
      
def refresh_Preview_MRA(s, event):
  # This function will populate the Matter Risk Assessment Preview datagrid
  #MessageBox.Show("Start - getting group ID", "Refreshing list (datagrid of questions)")
  if dg_FeeEarners.SelectedIndex == -1:
    tb_OurRef.Text = ""
    tb_MRA_Name.Text =  "NOTHING SELECTED"
    lbl_MRA_OV_ID.Content = "-1"
    tb_FE_Name.Text = ""
    #btn_ApproveHR.IsEnabled = False
    stk_SelectedMRAHeader_NoData.Visibility = Visibility.Visible
    stk_SelectedMRAHeader.Visibility = Visibility.Collapsed
    return

  if dg_FeeEarners.SelectedIndex > -1:
    tb_OurRef.Text = dg_FeeEarners.SelectedItem['OurRef']
    tb_MRA_Name.Text =  dg_FeeEarners.SelectedItem['MRAName']
    lbl_MRA_OV_ID.Content = dg_FeeEarners.SelectedItem['OV_ID']
    tb_FE_Name.Text = dg_FeeEarners.SelectedItem['FEName']
    #btn_ApproveHR.IsEnabled = True
    stk_SelectedMRAHeader_NoData.Visibility = Visibility.Collapsed
    stk_SelectedMRAHeader.Visibility = Visibility.Visible
    ov_ID = dg_FeeEarners.SelectedItem['OV_ID']

  #MessageBox.Show("Genating SQL...", "Refreshing list (datagrid of questions)")
  mySQL = """SELECT '0-QGroup' = QGs.Name, 
                    '1-DisplayOrder' = MRAD.DisplayOrder, 
                    '2-QText' = TQs.QuestionText, 
                    '3-Answer' = CASE WHEN MRAD.AnswerListToUse = '(TextBox)' THEN MRAD.tbAnswerText ELSE (SELECT AnswerText FROM Usr_MRA_TemplateAs TAs WHERE TAs.AnswerID = MRAD.SelectedAnswerID AND TAs.QuestionID = MRAD.QuestionID) END, 
                    '4-Notes' = MRAD.Notes, 
                    '5-EmailComment' = MRAD.EmailComment 
            FROM Usr_MRA_Detail MRAD 
              INNER JOIN Usr_MRA_TemplateQs TQs ON MRAD.QuestionID = TQs.QuestionID	
              INNER JOIN Usr_MRA_QGroups QGs ON MRAD.QGroupID = QGs.ID 
            WHERE MRAD.OV_ID = {0} 
            ORDER BY MRAD.QGroupID, MRAD.DisplayOrder 
  """.format(ov_ID)

  #mySQL = "SELECT '0-QGroup' = QGs.Name, '1-DisplayOrder' = MRAD.DisplayOrder, '2-QText' = TQs.QuestionText, "
  #mySQL += "'3-Answer' = CASE WHEN MRAD.AnswerListToUse = '(TextBox)' THEN MRAD.tbAnswerText ELSE (SELECT AnswerText FROM Usr_MRA_TemplateAs TAs WHERE TAs.AnswerID = MRAD.SelectedAnswerID) END, "
  #mySQL += "'4-Notes' = MRAD.Notes, '5-EmailComment' = MRAD.EmailComment "
  #mySQL += "FROM Usr_MRA_Detail MRAD "
  #mySQL += "INNER JOIN Usr_MRA_TemplateQs TQs ON MRAD.QuestionID = TQs.QuestionID	"
  #mySQL += "INNER JOIN Usr_MRA_QGroups QGs ON MRAD.QGroupID = QGs.ID "
  #mySQL += "WHERE MRAD.OV_ID = " + str(ov_ID) + " "
  #mySQL += "ORDER BY MRAD.QGroupID, MRAD.DisplayOrder "

  _tikitDbAccess.Open(mySQL)
  myItems = []
  
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          iQGroup = '' if dr.IsDBNull(0) else dr.GetString(0)
          iDO = 0 if dr.IsDBNull(1) else dr.GetValue(1)
          iQText = '-' if dr.IsDBNull(2) else dr.GetString(2)
          iAnswer = '' if dr.IsDBNull(3) else dr.GetString(3)
          iNotes = '' if dr.IsDBNull(4) else dr.GetString(4)
          iEC = '' if dr.IsDBNull(5) else dr.GetString(5)
          
          myItems.append(review_MRA(iDO, iQText, iAnswer, iNotes, iEC, iQGroup))
      
    dr.Close()
  _tikitDbAccess.Close()
  
  #MessageBox.Show("Putting list items into Datagrid", "Refreshing list (datagrid of questions)")
  #dg_MRAReview.ItemsSource = myItems   # (original - without 'Grouping')
  tmpC = ListCollectionView(myItems)
  tmpC.GroupDescriptions.Add(PropertyGroupDescription("QGroup"))
  dg_MRAReview.ItemsSource = tmpC
  
  
  if dg_MRAReview.Items.Count == 0:
    tb_OurRef.Text = ""
    tb_MRA_Name.Text =  "NOTHING SELECTED"
    lbl_MRA_OV_ID.Content = "-1"
    tb_FE_Name.Text = ""
    stk_SelectedMRAHeader_NoData.Visibility = Visibility.Visible
    stk_SelectedMRAHeader.Visibility = Visibility.Collapsed
    dg_MRAReview.Visibility = Visibility.Collapsed
  else:
    stk_SelectedMRAHeader_NoData.Visibility = Visibility.Collapsed
    stk_SelectedMRAHeader.Visibility = Visibility.Visible
    dg_MRAReview.Visibility = Visibility.Visible
  
  return


def GroupItems_Preview_SelectionChanged(s, event):
  refresh_Preview_MRA(s, event)
  dg_MRAReview.SelectedIndex = 0
  return
  

  
# # # #   END OF:   P R E V I E W   M A T T E R   R I S K   A S S E S S M E N T   TAB # # # #



def get_FullEntityRef(shortRef): 
  
  if len(shortRef) < 4:
    myFinalString = shortRef
  else:
    leftPart = shortRef[0:3] 
    rightPart = shortRef[3:7]
    noZerosToAdd = 15 - len(shortRef)
    myZeros = ''
    
    for x in range(noZerosToAdd):
      myZeros += '0'
    
    # combine text elements to create full length ref
    myFullLenString = leftPart + myZeros + rightPart
    
    # now need to actually check if Entity exists - following returns 0 if not valid entity, else returns number of entities with that ref - should only be one)
    countOfEntities = 0
    countOfEntities = _tikitResolver.Resolve("[SQL: SELECT COUNT(Code) FROM Entities WHERE Code = '{0}']".format(myFullLenString))
    # if above count is zero, we return the short ref so other functions return 'error' otherwise we provide the full length code
    if int(countOfEntities) == 0:
      myFinalString = shortRef
    else:
      myFinalString = myFullLenString
 
    #MessageBox.Show("GetFullLenEntityRef - Input: " + str(shortRef) + "\nOutput: " + myFinalString + "\nCount of Entities: " + str(countOfEntities))
  return myFinalString


def runSQL(codeToRun = "", showError = False, errorMsgText = "", errorMsgTitle = "", apostropheHandle = 0):
  # Traditionally, we used to use _tikitResolver.Resolve() as-is, but have since found that it's better to wrap this within Python's 'try: except:' construct.
  # In order to minimise code, I made this reusable function to do so plus allow for a custom message to be displayed upon error.  See below for explanation of inputs/arguments:
  # codeToRun     = Full SQL of code to run. No need to wrap in '[SQL: code_Here]' as we can do that here
  # showError     = True / False. Indicates whether or not to display message upon error
  # errorMsgText  = Text to display in the body of the message box upon error (note: actual SQL will automatically be included, so no need to re-supply that)
  # errorMsgTitle = Text to display in the title bar of the message box upon error
  # apostropheHandle = Toggle to escape apostrophes for the returned values
  
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
  
  # try to execute the SQL
  try:
    tmpValue = _tikitResolver.Resolve(fCodeToRun)
    if apostropheHandle == 1:
      tmpValue = tmpValue.replace("'", "''")
    return tmpValue
  except:
    # there was an error... check to see if opted to show message or not...
    return errorReturnValue
    if showError == True:
      MessageBox.Show(str(errorMsgText) + "\nSQL used:\n" + str(codeToRun), errorMsgTitle)
    #if len(str(errorReturnValue)) > 0:
    #  return errorReturnValue
    #else: 
    return ''


def isUserAnApprovalUser(userToCheck):
  # This is a new function to replace the 'isActiveUserHOD()' function (from 7th August 2024) - now use the 'Locks' table to see if user has Key to XAML HOD screens
  
  HODusersSQL = "SELECT STRING_AGG(UserRef, ' | ') FROM Keys WHERE LockRef = ( SELECT Code FROM Locks WHERE Description = 'XAML_Screen_HOD_AccessOnly')"
  HODusers = runSQL(HODusersSQL)

  if userToCheck in HODusers:
    return True
  else:
    return False

###################################################################################################################################################

def Approve_Button_Clicked(s, event):
  if dg_FeeEarners.SelectedIndex == -1:
    return
  
  # get input variables
  ovID = lbl_MRA_OV_ID.Content
  entRef = dg_FeeEarners.SelectedItem['EntityRef']
  matNo = dg_FeeEarners.SelectedItem['MatterNo']
  errorCount = 0
  errorMessage = ""

  tmpOurRef = dg_FeeEarners.SelectedItem['OurRef']
  tmpEmailTo = runSQL("SELECT EMailExternal FROM Users WHERE Code = (SELECT FeeEarnerRef FROM Matters WHERE EntityRef = '{0}' AND Number = {1})".format(entRef, matNo), False, '', '')
  tmpEmailCC = runSQL("SELECT EMailExternal FROM Users WHERE Code = '{0}'".format(_tikitUser), False, '', '')
  tmpToUserName = runSQL("SELECT ISNULL(Forename, FullName) FROM Users WHERE Code = (SELECT FeeEarnerRef FROM Matters WHERE EntityRef = '{0}' AND Number = {1})".format(entRef,  matNo), False, '', '', 1)
  tmpMatDesc = runSQL("SELECT Description FROM Matters WHERE EntityRef = '{0}' AND Number = {1}".format(entRef, matNo), False, '', '', 1)
  tmpClName = runSQL("SELECT LegalName FROM Entities WHERE Code = '{0}'".format(entRef), False, '', '', 1)
  tmpAddtl1 = dg_FeeEarners.SelectedItem['MRAName']
  tmpAddtl2 = "High"
  

  # generate SQL to approve
  approveSQL = "UPDATE Usr_MRA_Overview SET ApprovedByHOD = 'Y' WHERE ID = {0}".format(ovID)
  try:
    _tikitResolver.Resolve("[SQL: {0}]".format(approveSQL))
  except:
    errorCount += 1
    errorMessage = " - couldn't mark the selected item as approved\n" + str(approveSQL)
  
  # get SQL to Unlock matter
  #unlockCode = "EXEC TW_LockHandler '" + entRef + "', " + str(matNo) + ", 'LockedByRiskDept', 'UnLock'"
  lockID = runSQL("SELECT Code FROM Locks WHERE Description = 'LockedByRiskDept'", False, '', '')
  countMatterLocksSQL = "SELECT COUNT(EntityRef) FROM EntityMatterLocks WHERE EntityRef = '{0}' AND MatterNo = {1} AND LockID = {2}".format(entRef, matNo, lockID)
  countMatterLocks = runSQL(countMatterLocksSQL, False, '', '')
  
  if int(countMatterLocks) != 0:
    unlockCode = "DELETE FROM EntityMatterLocks WHERE EntityRef = '{0}' AND MatterNo = {1} AND LockID = {2}".format(entRef, matNo, lockID)
    try:
      _tikitResolver.Resolve("[SQL: {0}]".format(unlockCode))
    except:
      errorCount += 1
      errorMessage = " - couldn't unlock the selected matter\n" + str(unlockCode)

  # WE ALSO NEED TO TRIGGER THE TASK CENTRE TASK TO NOTIFY THE FE THAT MATTER HAS BEEN UNLOCKED
  tc_Trigger = """INSERT INTO Usr_MRA_Events (Date, UserRef, ActionTrigger, OV_ID, EmailTo, EmailCC, ToUserName, OurRef, MatterDesc, ClientName, Addtl1, Addtl2, FullEntityRef, Matter_No)
                  VALUES(GETDATE(), '{0}', 'HOD_Approved_MRA', {1}, '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', {11})
    """.format(_tikitUser, ovID, tmpEmailTo, tmpEmailCC, tmpToUserName, tmpOurRef, tmpMatDesc, tmpClName, tmpAddtl1, tmpAddtl2, entRef, matNo)

  #tc_Trigger = "INSERT INTO Usr_MRA_Events (Date, UserRef, ActionTrigger, OV_ID, EmailTo, EmailCC, ToUserName, OurRef, MatterDesc, ClientName, Addtl1, Addtl2) "
  #tc_Trigger += "VALUES(GETDATE(), '" + _tikitUser + "', 'HOD_Approved_MRA', " + str(ovID) + ", '" + tmpEmailTo + "', '" + tmpEmailCC + "', "
  #tc_Trigger += "'" + tmpToUserName + "', '" + str(tmpOurRef) + "', '" + str(tmpMatDesc) + "', '" + str(tmpClName) + "', '" + str(tmpAddtl1) + "', '" + str(tmpAddtl2) + "')"
  try:
    _tikitResolver.Resolve("[SQL: {0}]".format(tc_Trigger))
  except:
    errorCount += 1
    errorMessage = " - couldn't send the 'HOD Approved' Task Centre confirmation email to FE\n" + str(tc_Trigger)  
  
  if errorCount > 0:
    MessageBox.Show("The following error(s) were encountered:\n" + errorMessage + "\n\nPlease screenshot this message and send to IT.Support@thackraywilliams.com to investigate", "Error: Approve High-Risk Matter...")
  else:
    createNewMRA_BasedOnCurrent(s, event)
    MessageBox.Show("Successfully Approved and Unlocked the selected matter - Fee Earners list will now refresh", "Approve High-Risk Matter...")
    populate_FeeEarnersList(s, event)
  return
  

def createNewMRA_BasedOnCurrent(s, event):
  # this function will duplicate the active MRA
  
  # get input variables
  idItemToCopy = dg_FeeEarners.SelectedItem['OV_ID']
  nameToCopy = dg_FeeEarners.SelectedItem['MRAName']
  finalName = "{0} (copy of {1})".format(nameToCopy, idItemToCopy)
  finalName = finalName.replace("'", "''")
  mra_Expiry = getSQLDate(dg_FeeEarners.SelectedItem['MRA Expiry'])
  entRef = dg_FeeEarners.SelectedItem['EntityRef']
  matNo = dg_FeeEarners.SelectedItem['MatterNo']
  
  # generate SQL to copy high-level (Overview)
  insertOV_SQL = """INSERT INTO Usr_MRA_Overview (EntityRef, MatterNo, TypeID, ExpiryDate, LocalName, Score, RiskRating, ApprovedByHOD, DateAdded)
                    SELECT '{0}', {1}, TypeID, DATEADD(WEEK, 4, '{2}'), '{3}', Score, RiskRating, 'N', GETDATE()
                    FROM Usr_MRA_Overview WHERE ID = {4}""".format(entRef, matNo, mra_Expiry, finalName, idItemToCopy) 

  #insertOV_SQL = "INSERT INTO Usr_MRA_Overview (EntityRef, MatterNo, TypeID, ExpiryDate, LocalName, Score, RiskRating, ApprovedByHOD, DateAdded) "
  #insertOV_SQL += "SELECT '" + entRef + "', " + str(matNo) + ", TypeID, DATEADD(WEEK, 4, '" + str(mra_Expiry) + "'), '" + finalName + "', "
  #insertOV_SQL += "Score, RiskRating, 'N', GETDATE() FROM Usr_MRA_Overview WHERE ID = " + str(idItemToCopy)
  
  try:
    _tikitResolver.Resolve("[SQL: {0}]".format(insertOV_SQL))
  except:
    MessageBox.Show("There was an error creating a new MRA, using SQL:\n" + str(insertOV_SQL), "Error: Duplicate selected item...")
    return
    
  # now get row ID of items added
  rowID = runSQL("SELECT TOP 1 ID FROM Usr_MRA_Overview WHERE LocalName = '{0}' AND EntityRef = '{1}' AND MatterNo = {2} ORDER BY DateAdded DESC".format(finalName, entRef, matNo), False, "", "")
  #rowID = runSQL("SELECT TOP 1 ID FROM Usr_MRA_Overview WHERE LocalName = '" + finalName + "' AND EntityRef = '" + entRef + "' AND MatterNo = " + str(matNo) + " ORDER BY DateAdded DESC", False, "", "")

  if int(rowID) > 0:
    insertQ_SQL = """INSERT INTO Usr_MRA_Detail (EntityRef, MatterNo, OV_ID, QuestionID, AnswerListToUse, SelectedAnswerID, CurrentAnswerScore, DisplayOrder, QGroupID, CorrActionID)
                    SELECT '{0}', {1}, {2}, QuestionID, AnswerListToUse, SelectedAnswerID, CurrentAnswerScore, DisplayOrder, QGroupID, Null
                    FROM Usr_MRA_Detail WHERE OV_ID = {3}""".format(entRef, matNo, rowID, idItemToCopy)

    #insertQ_SQL = "INSERT INTO Usr_MRA_Detail (EntityRef, MatterNo, OV_ID, QuestionID, AnswerListToUse, SelectedAnswerID, CurrentAnswerScore, DisplayOrder, QGroupID, CorrActionID) "
    #insertQ_SQL += "SELECT '" + entRef + "', " + str(matNo) + ", " + str(rowID) + ", QuestionID, AnswerListToUse, SelectedAnswerID, "
    #insertQ_SQL += "CurrentAnswerScore, DisplayOrder, QGroupID, Null FROM Usr_MRA_Detail WHERE OV_ID = " + str(idItemToCopy)
    
    try:
      _tikitResolver.Resolve("[SQL: {0}]".format(insertQ_SQL))
    except:
      MessageBox.Show("An error occurred copying the Questions, using SQL:\n" + str(insertQ_SQL), "Error: Duplicate selected item - Copying Questions...")
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


class Users(object):
  def __init__(self, myCode, myName):
    self.Code = myCode
    self.Name = myName
    return
    
  def __getitem__(self, index):
    if index == 'Code':
      return self.Code
    elif index == 'Name':
      return self.Name
      
def populateUsersList(s, event):
  userSQL = "SELECT Code, FullName FROM Users WHERE UserStatus = 0 AND Locked = 0"
  
  _tikitDbAccess.Open(userSQL)
  uItem = []
  
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        iCode = '' if dr.IsDBNull(0) else dr.GetString(0)  
        iName = '' if dr.IsDBNull(1) else dr.GetString(1)  
        uItem.append(Users(iCode, iName))
    dr.Close()
  _tikitDbAccess.Close
  
  cbo_User.ItemsSource = uItem
  return


def setCurrentUser(s, event):
  pCount = -1
  for x in cbo_User.Items:
    pCount += 1 
    if x.Code == _tikitUser:
      cbo_User.SelectedIndex = pCount
      break
  return

]]>
    </Init>
    <Loaded>
      <![CDATA[
#Define controls that will be used in all of the code

stk_CurrentUser = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'stk_CurrentUser')
cbo_User = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cbo_User')
#cbo_User.DropDownClosed += populate_FeeEarnersList
cbo_User.SelectionChanged += populate_FeeEarnersList
lbl_UserStatus = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_UserStatus')

dg_FeeEarners = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_FeeEarners')
dg_FeeEarners.SelectionChanged += refresh_Preview_MRA

tb_OurRef = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_OurRef')
tb_MRA_Name = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MRA_Name')
tb_FE_Name = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_FE_Name')
lbl_MRA_OV_ID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRA_OV_ID')
btn_ApproveHR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_ApproveHR')
btn_ApproveHR.Click += Approve_Button_Clicked


#tb_NoMRA_PreviewQs = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_NoMRA_PreviewQs')
dg_MRAReview = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_MRAReview')
#dg_MRAReview.SelectionChanged += MRA_Preview_SelectionChanged
tb_NoFEs = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_NoFEs')

stk_SelectedMRAHeader_NoData = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'stk_SelectedMRAHeader_NoData')
stk_SelectedMRAHeader = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'stk_SelectedMRAHeader')



# Define Actions and on load events
myOnLoadEvent(_tikitSender, 'onLoad')

]]>
    </Loaded>
  </RiskPractice>

</tfb>