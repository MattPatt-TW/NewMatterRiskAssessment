<?xml version="1.0" encoding="UTF-8"?>
<tfb>
  <Risk_HODApproval_v2>
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
  
  populateUsersList()
  setCurrentUser()
  
  whoCanChangeCurrentUser = ['LD1', 'MP', 'NB', 'AH1']
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
  def __init__(self, myFEName, myBranch, myOurRef, myMRAname, myMRAexp, myMRAID, myEntRef, myMatNo, myType, myClName, myMatDesc, myFEEmail):
    self.feName = myFEName
    self.feBranch = myBranch
    self.ourRef = myOurRef
    self.mraName = myMRAname
    self.mraExpiryDate = myMRAexp
    self.mraID = myMRAID
    self.entRef = myEntRef
    self.matNo = myMatNo
    self.AType = myType
    self.clName = myClName
    self.matDesc = myMatDesc
    self.feEmail = myFEEmail
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
    elif index == 'mraID':
      return self.mraID
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
    elif index == 'FEEmail':
      return self.feEmail


def populate_FeeEarnersList(s, event): 

  # FOLLOWING CODE IS DIRECTLY FROM WIP SCREEN
  currentUser = cbo_User.SelectedItem['Code']
  #user_Dept = runSQL("SELECT Department FROM Users WHERE Code = '" + currentUser + "'", False, "", "")
  #userIsHOD = isUserAnApprovalUser(userToCheck = currentUser)    # can't see this is being used for anything
  myFEitems = []
  
  mySQL = """SELECT '0-FeeEarner' = U.FullName, 
                    '1-Branch' = B.Description, 
                    '2-OurRef' = CONCAT(LEFT(MH.EntityRef, 3), RIGHT(MH.EntityRef, 4), '/', CONVERT(nvarchar, MH.MatterNo, 1)), 
                    '3-MRA Name' = MH.Name, 
                    '4-Expiry Date' = MH.ExpiryDate, 
                    '5-mraID' = MH.mraID, 
                    '6-EntityRef' = MH.EntityRef, 
                    '7-MatterNo' = MH.MatterNo, 
                    '8-CaseType' = CT.Description, 
                    '9-Type' = CASE WHEN '{0}' IN (SELECT HA.UserCode FROM Usr_HODapprovals HA WHERE HA.FeeEarnerCode = U.Code AND HA.Type = 'Main') THEN '01) Main' 
                                    WHEN '{0}' IN (SELECT HA.UserCode FROM Usr_HODapprovals HA WHERE HA.FeeEarnerCode = U.Code AND HA.Type = 'Holiday Cover') THEN '02) Holiday Cover' ELSE '' END, 
                    '10-MatterDesc' = M.Description, 
                    '11-EntName' = E.LegalName, 
                    '12-FEEmail' = U.EMailExternal, 
        FROM Usr_MRAv2_MatterHeader MH
          INNER JOIN Matters M ON MH.EntityRef = M.EntityRef AND MH.MatterNo = M.Number 
          LEFT OUTER JOIN Users U ON M.FeeEarnerRef = U.Code 
          LEFT OUTER JOIN Branches B ON U.UsualOffice = B.Code
          INNER JOIN CaseTypes CT ON M.CaseTypeRef = CT.Code
          INNER JOIN CaseTypeGroups CTG ON CT.CaseTypeGroupRef = CTG.ID
          LEFT OUTER JOIN Usr_Approvals App ON U.Code = App.UserCode
          LEFT OUTER JOIN Entities E ON MH.EntityRef = E.Code
        WHERE MH.ApprovedByHOD = 'N' AND MH.RiskRating = 3 AND MH.Status = 'Complete'
          AND U.Code IN (SELECT FeeEarnerCode FROM Usr_HODapprovals WHERE UserCode = '{0}')
        ORDER BY '9-Type', '0-FeeEarner', '4-Expiry Date' """.format(currentUser)

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
          imraID = 0 if dr.IsDBNull(5) else dr.GetValue(5)
          iEntRef = 0 if dr.IsDBNull(6) else dr.GetString(6)
          iMatNo = 0 if dr.IsDBNull(7) else dr.GetValue(7)
          iType = '' if dr.IsDBNull(9) else dr.GetString(9)
          iMatDesc = '' if dr.IsDBNull(10) else dr.GetString(10)
          iClName = '' if dr.IsDBNull(11) else dr.GetString(11)
          iFEEmail = '' if dr.IsDBNull(12) else dr.GetString(12)
          
          myFEitems.append(UsersList(iFE, iBranch, iOurRef, iMRAName, iMRAExpiry, imraID, iEntRef, iMatNo, iType, iClName, iMatDesc, iFEEmail))
      
    dr.Close()
  _tikitDbAccess.Close()
  
  # now group Fee Earners by Branch and set THAT object as the ItemsSource...
  tmpC = ListCollectionView(myFEitems)
  #tmpC.GroupDescriptions.Add(PropertyGroupDescription("feBranch"))
  tmpC.GroupDescriptions.Add(PropertyGroupDescription("AType"))
  tmpC.GroupDescriptions.Add(PropertyGroupDescription("feName"))
  
  dg_FeeEarners.ItemsSource = tmpC
  
  # if nothing in list
  if dg_FeeEarners.Items.Count == 0:
    # show 'no data' text and hide the rest
    tb_NoFEs.Visibility = Visibility.Visible
    stk_SelectedMRAHeader_NoData.Visibility = Visibility.Visible
    
    stk_SelectedMRAHeader.Visibility = Visibility.Collapsed
    dg_MRAReview.Visibility = Visibility.Collapsed
    dg_FeeEarners.Visibility = Visibility.Collapsed
  else:
    # there ARE items in the list, so show datagrid and hide 'no data' text
    tb_NoFEs.Visibility = Visibility.Collapsed
    stk_SelectedMRAHeader_NoData.Visibility = Visibility.Collapsed
    
    stk_SelectedMRAHeader.Visibility = Visibility.Visible
    dg_MRAReview.Visibility = Visibility.Visible
    dg_FeeEarners.Visibility = Visibility.Visible
  return
  
  
  

class review_MRA(object):
  def __init__(self, myOrder, myQuestion, myAnswerText, myFEnotes, myEC, myGroup, myScore):
    self.DOrder = myOrder
    self.QuestionText = myQuestion
    self.LUP_AnswerText = myAnswerText
    self.FENotes = myFEnotes
    self.LUP_EmailComment = myEC
    self.QGroup = myGroup
    self.Score = myScore
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
    elif index == 'Score':
      return self.Score
    else:
      return ''
      
def refresh_Preview_MRA(s, event):
  # This function will populate the Matter Risk Assessment Preview datagrid
  #MessageBox.Show("Start - getting group ID", "Refreshing list (datagrid of questions)")
  if dg_FeeEarners.SelectedIndex == -1:
    tb_OurRef.Text = ""
    tb_MRA_Name.Text =  "NOTHING SELECTED"
    tb_MRA_ID.Text = "-1"
    tb_FE_Name.Text = ""
    #btn_ApproveHR.IsEnabled = False
    stk_SelectedMRAHeader_NoData.Visibility = Visibility.Visible
    stk_SelectedMRAHeader.Visibility = Visibility.Collapsed
    return

  # otherwise an item is selected, so populate controls with data from selected DG item
  tb_OurRef.Text = dg_FeeEarners.SelectedItem['OurRef']
  tb_MRA_Name.Text =  dg_FeeEarners.SelectedItem['MRAName']
  tb_MRA_ID.Text = dg_FeeEarners.SelectedItem['mraID']
  tb_ClName.Text = dg_FeeEarners.SelectedItem['ClientName']
  tb_MatDesc.Text = dg_FeeEarners.SelectedItem['MatDesc']
  tb_FE_Name.Text = dg_FeeEarners.SelectedItem['FEName']
  tb_FE_Email.Text = dg_FeeEarners.SelectedItem['FEEmail']

  #btn_ApproveHR.IsEnabled = True
  stk_SelectedMRAHeader_NoData.Visibility = Visibility.Collapsed
  stk_SelectedMRAHeader.Visibility = Visibility.Visible
  mraID = dg_FeeEarners.SelectedItem['mraID']

  #MessageBox.Show("Genating SQL...", "Refreshing list (datagrid of questions)")
  mySQL = """SELECT '0-QGroup' = MD.QuestionGroup, 
                    '1-DisplayOrder' = MD.DisplayOrder, 
                    '2-QText' = TQs.QuestionText, 
                    '3-Answer' = TAs.AnswerText, 
                    '4-Notes' = MD.Comments, 
                    '5-EmailComment' = MD.EmailComment, 
                    '6-Score' = MD.Score
             FROM Usr_MRAv2_MatterDetails MD 
               INNER JOIN Usr_MRAv2_Question TQs ON MD.QuestionID = TQs.QuestionID
               INNER JOIN Usr_MRAv2_Answer TAs ON MD.AnswerID = TAs.AnswerID
             WHERE MD.mraID = {0} 
             ORDER BY MD.QuestionGroup, MD.DisplayOrder """.format(mraID)

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
          iScore = 0 if dr.IsDBNull(6) else dr.GetValue(6)
          
          myItems.append(review_MRA(iDO, iQText, iAnswer, iNotes, iEC, iQGroup, iScore))
      
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
    tb_MRA_ID.Text = "-1"
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
  mraID = tb_MRA_ID.Text
  entRef = dg_FeeEarners.SelectedItem['EntityRef']
  matNo = dg_FeeEarners.SelectedItem['MatterNo']
  errorCount = 0
  errorMessage = ""

  tmpOurRef = dg_FeeEarners.SelectedItem['OurRef']
  #! 26/02/2026 - eek... this could do with a re-write, far too many individual SQL calls to get the various pieces of information we want to include 
  #! in the Task Centre email - would be better to get all info in one go at the start of this function and store in variables, then reference those variables when building the email (rather than having multiple calls to the database throughout this function)
  # May be worth adding into initial FE list, to get their details as applicable and save in DataGrid/object
  # Then when selection is changed, and we populate right-hand side, might be worth adding text boxes in the header area so it's
  # clearly visible to HOD if Email is missing etc.
  # Then, we just reference those text boxes/variables here

  tmpEmailTo = tb_FE_Email.Text
  tmpEmailCC = cbo_User.SelectedItem['Email']
  tmpToUserName = tb_FE_Name.Text
  tmpMatDesc = tb_MatDesc.Text
  tmpClName = tb_ClName.Text
  tmpAddtl1 = tb_MRA_Name.Text
  tmpAddtl2 = "High"
  

  # generate SQL to approve
  approveSQL = """UPDATE Usr_MRAv2_MatterHeader SET ApprovedByHOD = 'Y' 
                  WHERE mraID = {mraID} AND EntityRef = '{entRef}' AND MatterNo = {matNo}""".format(mraID=mraID, entRef=entRef, matNo=matNo)
  try:
    _tikitResolver.Resolve("[SQL: {0}]".format(approveSQL))
  except:
    errorCount += 1
    errorMessage = " - couldn't mark the selected item as approved\n" + str(approveSQL)
  
  # get SQL to Unlock matter
  #unlockCode = "EXEC TW_LockHandler '" + myEntRef + "', " + str(myMatNo) + ", 'LockedByRiskDept', 'UnLock'"
  unlockCode = "EXEC TW_LockHandler '{entRef}', {matNo}, 'LockedByRiskDept', 'UnLock'".format(entRef=entRef, matNo=matNo)
  # run unlock code
  runSQL(codeToRun=unlockCode, showError=True, 
         errorMsgText="There was an error unlocking the matter, after approval. Please check the matter is unlocked and if not, unlock manually using the following SQL:\n{0}".format(unlockCode), errorMsgTitle="Error: Unlocking Matter after Approval...")

  # WE ALSO NEED TO TRIGGER THE TASK CENTRE TASK TO NOTIFY THE FE THAT MATTER HAS BEEN UNLOCKED
  tc_Trigger = """INSERT INTO Usr_MRA_Events (Date, UserRef, ActionTrigger, OV_ID, EmailTo, EmailCC, ToUserName, 
                                              OurRef, MatterDesc, ClientName, Addtl1, Addtl2, FullEntityRef, Matter_No)
                  VALUES(GETDATE(), '{user}', 'HOD_Approved_MRA', {mraID}, '{emailTo}', '{emailCC}', '{toUserName}', 
                        '{ourRef}', '{matDesc}', '{clientName}', '{addtl1}', '{addtl2}', '{fullEntRef}', {matNo})
    """.format(user=_tikitUser, mraID=mraID, emailTo=tmpEmailTo, emailCC=tmpEmailCC, toUserName=tmpToUserName, 
               ourRef=tmpOurRef, matDesc=tmpMatDesc, clientName=tmpClName, addtl1=tmpAddtl1, addtl2=tmpAddtl2, fullEntRef=entRef, matNo=matNo)

  try:
    _tikitResolver.Resolve("[SQL: {0}]".format(tc_Trigger))
  except:
    errorCount += 1
    errorMessage = " - couldn't send the 'HOD Approved' Task Centre confirmation email to FE\n" + str(tc_Trigger)  
  
  if errorCount > 0:
    MessageBox.Show("The following error(s) were encountered:\n" + errorMessage + "\n\nPlease screenshot this message and send to IT.Support@thackraywilliams.com to investigate", "Error: Approve High-Risk Matter...")
  else:
    createNewMRA_BasedOnCurrent(idItemToCopy=mraID, entRef=entRef, matNo=matNo)
    MessageBox.Show("Successfully Approved and Unlocked the selected matter - Fee Earners list will now refresh", "Approve High-Risk Matter...")
    populate_FeeEarnersList(s, event)
  return
  

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
  def __init__(self, myCode, myName, myEmail):
    self.Code = myCode
    self.Name = myName
    self.Email = myEmail
    return
    
  def __getitem__(self, index):
    if index == 'Code':
      return self.Code
    elif index == 'Name':
      return self.Name
    elif index == 'Email':
      return self.Email
      
def populateUsersList():
  userSQL = "SELECT Code, FullName, EMailExternal FROM Users WHERE UserStatus = 0 AND Locked = 0"
  
  _tikitDbAccess.Open(userSQL)
  uItem = []
  
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        iCode = '' if dr.IsDBNull(0) else dr.GetString(0)  
        iName = '' if dr.IsDBNull(1) else dr.GetString(1)  
        iEmail = '' if dr.IsDBNull(2) else dr.GetString(2)
        uItem.append(Users(iCode, iName, iEmail))
    dr.Close()
  _tikitDbAccess.Close()
  
  cbo_User.ItemsSource = uItem
  return


def setCurrentUser():
  pCount = -1
  for x in cbo_User.Items:
    pCount += 1 
    if x.Code == _tikitUser:
      cbo_User.SelectedIndex = pCount
      break
  return


def get_NextMRAFR_NumberForMatter(ovID = 0, entityRef = '', matterNo = ''):
  # This new function was added 20/05/2025 as there are a couple of occurences where we need to get the next
  # MRA/FR number for a given TypeID (testing against current Entity/Matter record)

  # if passed ID is empty, exit and alert user
  if ovID == 0:
    MessageBox.Show("You need to pass an ID to this function!", "Error: get_NextMRAFR_NumberForMatter...")
    return 0
  
  # else we carry on abd get the TypeID
  tmpTypeID = runSQL("SELECT TypeID FROM Usr_MRA_Overview WHERE ID = {0} AND EntityRef = '{1}' AND MatterNo = {2}".format(ovID, entityRef, matterNo), False, '', '')

  NextNum_sql = """[SQL: SELECT COUNT(TypeID) + 1 FROM Usr_MRA_Overview MRAO 
                         WHERE MRAO.EntityRef = '{0}' AND MRAO.MatterNo = {1} 
                          AND TypeID = {2}]""".format(entityRef, matterNo, tmpTypeID)
  NextNum = runSQL(NextNum_sql, False, '', '')
  return NextNum


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
tb_ClName = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_ClName')
tb_MatDesc = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MatDesc')
tb_FE_Email = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_FE_Email')
tb_FE_Name = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_FE_Name')
tb_MRA_ID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MRA_ID')
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
  </Risk_HODApproval_v2>

</tfb>