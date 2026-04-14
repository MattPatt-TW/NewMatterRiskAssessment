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

from System import DateTime, Double
from System.Diagnostics import Process
from System.Globalization import DateTimeStyles
from System.Collections.Generic import Dictionary
from System.Windows import Controls, Forms, LogicalTreeHelper, FrameworkElement, Window
from System.Windows import Data, UIElement, Visibility, Window
from System.Windows.Controls import Button, Canvas, GridView, GridViewColumn, ListView, Orientation, ScrollViewer
from System.Windows.Data import Binding, CollectionView, ListCollectionView, PropertyGroupDescription
from System.Windows.Forms import SelectionMode, MessageBox, MessageBoxButtons, DialogResult
from System.Windows.Input import KeyEventHandler
from System.Windows.Media import Brush, Brushes, VisualTreeHelper

## GLOBAL VARIABLES ##
preview_MRA = []    # To temp store table for previewing Matter Risk Assessment
previewFR = []      # To temp store table for previewing File Review
lockID_LockedByRiskDept = 0   # To store the LockID for the 'LockedByRiskDept' lock, which we use to identify locked matters in our SQL queries (set on load)

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
  set_lockID_for_LockedByRiskDept()
  form_resize_to_host()

  #MessageBox.Show("Hello world! - OnLoad Finished")
  return

def form_resize_to_host():
  root = Risk_HODApproval_v2

  host_sv = find_ancestor_of_type(root, ScrollViewer)
  if host_sv is not None:
    host_sv.SizeChanged += resize_to_host
    host_sv.ScrollChanged += resize_to_host

  parent_fe = get_nearest_parent_framework_element(root)
  if parent_fe is not None:
      parent_fe.SizeChanged += resize_to_host  

  #host_window = find_ancestor_of_type(root, Window)
  #if host_window is not None:
  #    host_window.SizeChanged += resize_to_host
  #    host_window.StateChanged += resize_to_host

  resize_to_host()


def set_lockID_for_LockedByRiskDept():
  global lockID_LockedByRiskDept

  if lockID_LockedByRiskDept == 0:
    lockID_SQL = """SELECT CASE WHEN EXISTS(SELECT Code FROM Locks WHERE Description = 'LockedByRiskDept') THEN 
                    (SELECT Code FROM Locks WHERE Description = 'LockedByRiskDept') ELSE 0 END """
    lockID_LockedByRiskDept = runSQL(lockID_SQL, False, '', '')
  
  if int(lockID_LockedByRiskDept) == 0:
    MessageBox.Show("There doesn't appear to be a Lock setup in the name of 'LockedByRiskDept', so cannot identify locked matters!", "Error: Setting LockID for LockedByRiskDept...")
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
  myFEitems = []
  
  mySQL = """;WITH HODAccess AS (
                  SELECT
                      HA.FeeEarnerCode,
                      MAX(CASE WHEN HA.Type = 'Main' THEN 1 ELSE 0 END) AS HasMain,
                      MAX(CASE WHEN HA.Type = 'Holiday Cover' THEN 1 ELSE 0 END) AS HasHoliday
                  FROM Usr_HODapprovals HA
                  WHERE HA.UserCode = '{0}'
                  GROUP BY HA.FeeEarnerCode
                )
              SELECT
                  U.FullName AS [0-FeeEarner],
                  B.Description AS [1-Branch],
                  CONCAT(E.ShortCode, '/', CONVERT(NVARCHAR(20), MH.MatterNo)) AS [2-OurRef],
                  MH.Name AS [3-MRA Name],
                  MH.ExpiryDate AS [4-Expiry Date],
                  MH.mraID AS [5-mraID],
                  MH.EntityRef AS [6-EntityRef],
                  MH.MatterNo AS [7-MatterNo],
                  CT.Description AS [8-CaseType],
                  CASE WHEN HA.HasMain = 1 THEN '01) Main'
                      WHEN HA.HasHoliday = 1 THEN '02) Holiday Cover'
                      ELSE ''
                  END AS [9-Type],
                  M.Description AS [10-MatterDesc],
                  E.LegalName AS [11-EntName],
                  U.EMailExternal AS [12-FEEmail]
              FROM Usr_MRAv2_MatterHeader MH
                INNER JOIN Matters M    ON MH.EntityRef = M.EntityRef    AND MH.MatterNo = M.Number
                LEFT JOIN Users U       ON M.FeeEarnerRef = U.Code
                LEFT JOIN Branches B    ON U.UsualOffice = B.Code
                INNER JOIN CaseTypes CT ON M.CaseTypeRef = CT.Code
                LEFT JOIN Entities E    ON MH.EntityRef = E.Code
                INNER JOIN HODAccess HA ON HA.FeeEarnerCode = U.Code
              WHERE MH.ApprovedByHOD = 'N' AND MH.RiskRating = 3 AND MH.Status = 'Complete'
              ORDER BY [9-Type], [0-FeeEarner], [4-Expiry Date];""".format(currentUser)

  _tikitDbAccess.Open(mySQL)

  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          iFE = '-' if dr.IsDBNull(0) else dr.GetString(0)
          iBranch = '-' if dr.IsDBNull(1) else dr.GetString(1)
          iOurRef = '' if dr.IsDBNull(2) else dr.GetString(2)
          iMRAName = '' if dr.IsDBNull(3) else dr.GetString(3)
          iMRAExpiry = 0 if dr.IsDBNull(4) else dr.GetValue(4)
          imraID = 0 if dr.IsDBNull(5) else dr.GetValue(5)
          iEntRef = '' if dr.IsDBNull(6) else dr.GetString(6)
          iMatNo = 0 if dr.IsDBNull(7) else dr.GetValue(7)
          iType = '' if dr.IsDBNull(9) else dr.GetString(9)
          iMatDesc = '' if dr.IsDBNull(10) else dr.GetString(10)
          iClName = '' if dr.IsDBNull(11) else dr.GetString(11)
          iFEEmail = '' if dr.IsDBNull(12) else dr.GetString(12)
          
          myFEitems.append(UsersList(myFEName=iFE, myBranch=iBranch, myOurRef=iOurRef, myMRAname=iMRAName, myMRAexp=iMRAExpiry, 
                                     myMRAID=imraID, myEntRef=iEntRef, myMatNo=iMatNo, myType=iType, myClName=iClName, myMatDesc=iMatDesc, 
                                     myFEEmail=iFEEmail))
      
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
  tb_OurRef.Text = str(dg_FeeEarners.SelectedItem['OurRef'])
  tb_MRA_Name.Text =  str(dg_FeeEarners.SelectedItem['MRAName'])
  tb_MRA_ID.Text = str(dg_FeeEarners.SelectedItem['mraID'])
  tb_ClName.Text = str(dg_FeeEarners.SelectedItem['ClientName'])
  tb_MatDesc.Text = str(dg_FeeEarners.SelectedItem['MatDesc'])
  tb_FE_Name.Text = str(dg_FeeEarners.SelectedItem['FEName'])
  tb_FE_Email.Text = str(dg_FeeEarners.SelectedItem['FEEmail'])

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
          
          myItems.append(review_MRA(myOrder=iDO, myQuestion=iQText, myAnswerText=iAnswer, myFEnotes=iNotes, myEC=iEC, myGroup=iQGroup, myScore=iScore))
      
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
  # This is a new function to replace the 'isActiveUserHOD()' function (from 7th August 2024) - now use the 'Locks' table to see if user has Key to XAML HOD screens
  
  HODusersSQL = "SELECT STRING_AGG(UserRef, ' | ') FROM Keys WHERE LockRef = ( SELECT Code FROM Locks WHERE Description = 'XAML_Screen_HOD_AccessOnly')"
  HODusers = runSQL(HODusersSQL)

  if userToCheck in HODusers:
    return True
  else:
    return False

###################################################################################################################################################
def Approve_Button_Clicked(s, event):
  # this is the modified version of the Approve button click function, which now calls a stored procedure to
  #  handle the approval, unlocking and duplication of the MRA in one go, rather than having multiple separate
  #  SQL calls from the code - this should make it more efficient, and also less prone to errors/failures part
  #  way through the process (e.g. if we approve the MRA but then fail to unlock the matter, it would be left
  #  in an inconsistent state where it's approved but still locked, which would require manual intervention to fix - by having all logic in one stored procedure, we can ensure that either all steps succeed, or if there's an error, none of the steps are applied, leaving data in a consistent state either way)

  if dg_FeeEarners.SelectedIndex == -1:
    return

  mraID = int(tb_MRA_ID.Text)
  entRef = str(dg_FeeEarners.SelectedItem['EntityRef'])
  matNo = int(dg_FeeEarners.SelectedItem['MatterNo'])

  tmpOurRef = str(dg_FeeEarners.SelectedItem['OurRef'])
  tmpEmailTo = str(tb_FE_Email.Text)
  tmpEmailCC = str(cbo_User.SelectedItem['Email'])
  tmpToUserName = str(tb_FE_Name.Text)
  tmpMatDesc = str(tb_MatDesc.Text)
  tmpClName = str(tb_ClName.Text)
  tmpAddtl1 = str(tb_MRA_Name.Text)
  tmpAddtl2 = "High"

  procSQL = """
    DECLARE @NewMRAID INT;

    EXEC dbo.TW_MRA_HODApprovesHighRiskMatter
        @mraID = {mraID},
        @EntityRef = '{entRef}',
        @MatterNo = {matNo},
        @ApprovedByUserRef = '{userRef}',
        @EmailTo = '{emailTo}',
        @EmailCC = '{emailCC}',
        @ToUserName = '{toUserName}',
        @OurRef = '{ourRef}',
        @MatterDesc = '{matDesc}',
        @ClientName = '{clientName}',
        @Addtl1 = '{addtl1}',
        @Addtl2 = '{addtl2}',
        @NewMRAID = @NewMRAID OUTPUT;

    SELECT @NewMRAID;
    """.format(
        mraID=mraID,
        entRef=entRef.replace("'", "''"),
        matNo=matNo,
        userRef=str(_tikitUser).replace("'", "''"),
        emailTo=tmpEmailTo.replace("'", "''"),
        emailCC=tmpEmailCC.replace("'", "''"),
        toUserName=tmpToUserName.replace("'", "''"),
        ourRef=tmpOurRef.replace("'", "''"),
        matDesc=tmpMatDesc.replace("'", "''"),
        clientName=tmpClName.replace("'", "''"),
        addtl1=tmpAddtl1.replace("'", "''"),
        addtl2=tmpAddtl2.replace("'", "''")
    )

  newMRAID = runSQL(codeToRun=procSQL,
                    useAlternativeResolver=True,
                    showError=True,
                    errorMsgText="There was an error approving/unlocking/duplicating the selected matter.",
                    errorMsgTitle="Error: Approve High-Risk Matter..."
    )

  if int(newMRAID) <= 0:
    return

  MessageBox.Show("Successfully Approved and Unlocked the selected matter - Fee Earners list will now refresh", "Approve High-Risk Matter...")
  populate_FeeEarnersList(s, event)


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



def find_ancestor_of_type(start_obj, target_type):
  current = start_obj
  while current is not None:
    current = VisualTreeHelper.GetParent(current)
    if current is not None and isinstance(current, target_type):
      return current
  return None

def get_nearest_parent_framework_element(start_obj):
  current = start_obj
  while current is not None:
    current = VisualTreeHelper.GetParent(current)
    if current is not None and isinstance(current, FrameworkElement):
      return current
  return None

def resize_to_host(s=None, event=None):
  root = Risk_HODApproval_v2
  host_sv = find_ancestor_of_type(root, ScrollViewer)
  target_height = 0

  if host_sv is not None and host_sv.ViewportHeight > 0:
    target_height = host_sv.ViewportHeight
  else:
    parent_fe = get_nearest_parent_framework_element(root)
    if parent_fe is not None and parent_fe.ActualHeight > 0:
      target_height = parent_fe.ActualHeight

  if target_height <= 0:
    return

  # allow for margins / host padding
  target_height -= 4
  if target_height < 100:
    return

  # Clear old constraints first, then apply fresh ones
  #root.ClearValue(FrameworkElement.HeightProperty)
  #root.ClearValue(FrameworkElement.MaxHeightProperty)

  root.Height = target_height
  root.MaxHeight = target_height
  root.UpdateLayout()


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

#################################################################################################
# code for the main forms' OK/Apply/Cancel buttons, meaning we can hide here too
# (useful for those screens where the XAML is used for multi-matters, and 'Apply' won't do anything, and technically 'OK' and 'Cancel' do same thing - close screen)
# Note: following 3 are needed to get a handle on main screen elements (which obviously are NOT on our XAML as they sit ABOVE it in the main Tikit form structure)
myScrollViewer = LogicalTreeHelper.GetParent(_tikitSender)
myDockPanel = LogicalTreeHelper.GetParent(myScrollViewer)
myGrid = LogicalTreeHelper.GetParent(myDockPanel)

tikitOK = LogicalTreeHelper.FindLogicalNode(myGrid, 'OK')
tikitApply = LogicalTreeHelper.FindLogicalNode(myGrid, 'Apply')
tikitCancel = LogicalTreeHelper.FindLogicalNode(myGrid, 'Cancel')
tikitContact = LogicalTreeHelper.FindLogicalNode(myGrid, 'Contacts')

tikitCancel.Content = 'Close'
tikitOK.Visibility = Visibility.Collapsed
tikitApply.Visibility = Visibility.Collapsed
tikitContact.Visibility = Visibility.Collapsed
#################################################################################################

Risk_HODApproval_v2 = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'Risk_HODApproval_v2')

# Define Actions and on load events
myOnLoadEvent(_tikitSender, 'onLoad')

]]>
    </Loaded>
  </Risk_HODApproval_v2>

</tfb>