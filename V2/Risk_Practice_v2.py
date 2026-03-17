<?xml version="1.0" encoding="UTF-8"?>
<tfb>
  <RiskPractice_v2>
    <Init>
      <![CDATA[
import clr
#from TWUtils import runSQL

clr.AddReference('mscorlib')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('System.Windows.Forms')

from datetime import datetime
from System import DateTime
from System.Diagnostics import Process
from System.Globalization import DateTimeStyles
from System.Collections.Generic import Dictionary
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import Controls, Forms, LogicalTreeHelper, RoutedEventHandler
from System.Windows import Data, UIElement, Visibility, Window
from System.Windows.Controls import Button, Canvas, GridView, GridViewColumn, ListView, Orientation
from System.Windows.Data import Binding, CollectionView, ListCollectionView, PropertyGroupDescription
from System.Windows.Forms import SelectionMode, MessageBox, MessageBoxButtons, DialogResult, MessageBoxIcon
from System.Windows.Input import KeyEventHandler
from System.Windows.Media import Brush, Brushes
import re

## GLOBAL VARIABLES ##
preview_MRA = []    # To temp store table for previewing Matter Risk Assessment
previewFR = []      # To temp store table for previewing File Review


# # # #   O N   L O A D   E V E N T   # # # #
def myOnLoadEvent(s, event):
  # populate drop-downs
  populate_FeeEarnersList(s, event)
  populate_DepartmentList(s, event)
  populate_GroupByCombo(s, event)
  
  #MessageBox.Show("DEBUG - POPULATING LISTS")
  # populate lists (DataGrids)
  refresh_ListOfLockedMatters()               # Locked Matters refresh
  refresh_FR_Templates(s, event)              # File Review (main overview / templates)
  refresh_FR_Department_Defaults(s, event)    # FR Default Template for Department
  refresh_FR_CaseType_Defaults(s, event)      # FR Default Template for Case Type
  refresh_AnswerListGroups(s, event)          # Answers list (shared for both MRA and FR)
  refresh_GroupItems(s, event)                # New 'Group' items (for MRA 'Section/Group')

  set_Visibility_ofAnswerItemsDG()  
  
  # wire up new 'New' button popup
  icTemplates.AddHandler(Button.ClickEvent,
                         RoutedEventHandler(TemplateButton_Click))

  # hide 'Edit Questions' and 'Preview' tabs
  ti_FR_Questions.Visibility = Visibility.Collapsed
  ti_FR_Preview.Visibility = Visibility.Collapsed
  
  #MessageBox.Show("Hello world! - OnLoad Finished")
  return

# # # #   L O C K E D   M A T T E R S   # # # #


class UnlockReason(object):
  def __init__(self, reason):
    self.Reason = reason
    return

def load_UnlockReasons():

  unlock_Reasons = [
        'Fee Earner Delay',
        'Fee Earner Illness',
        'File not fully open',
        'Client completing e-verification',
        'Wait on documentation/info',
        'Matter did not proceed',
        'Late HOD approval of High Risk',
        'Holiday',
        'Appointment booked in the future',
        'Test File',
        'Historic File/ Previous FE error'
        ]

  # now we'll iterate over this list and create a list of 'UnlockReason' objects, which we can then bind to the pop-up list (ItemsControl)
  unlock_Reasons_Objects = ObservableCollection[object]()
  for r in unlock_Reasons:
    unlock_Reasons_Objects.Add(UnlockReason(r))
  return unlock_Reasons_Objects


def btn_UnlockMatters2_Click(s, event):
  # ToggleButton behaviour: use IsChecked to decide open state
  is_open = bool(s.IsChecked)
  if is_open:
    # get reasons list
    icTemplates.ItemsSource = load_UnlockReasons()

    popTemplates.IsOpen = True
  else:
    popTemplates.IsOpen = False


def TemplateButton_Click(sender, args):
  # Because we used AddHandler on icTemplates, sender may be the ItemsControl.
  # args.OriginalSource should be the actual Button (or something inside it).

  btn = args.OriginalSource
  # Sometimes OriginalSource is a TextBlock inside the Button; walk up to Button
  while btn is not None and not isinstance(btn, Button):
    btn = getattr(btn, "TemplatedParent", None) or getattr(btn, "Parent", None)

  if btn is None:
    return

  opt = btn.Tag  # <-- TemplateOption
  reasonText = str(opt.Reason)  # This is the text of the button, which is the reason for unlocking
  sql_reasonText = reasonText.replace("'", "''")  # Escape single quotes for SQL
  # do unlock code passing in opt as the reason
  #MessageBox.Show("You clicked the button for ReasonText: {0}".format(sql_reasonText), "Unlock reason selected - debugging...")
  unlockMatter(withReason=sql_reasonText)

  # Close popup
  popTemplates.IsOpen = False


def CancelPopup_Click(sender, args):
  popTemplates.IsOpen = False


def popTemplates_Closed(sender, args):
  # Ensure the toggle button pops back up
  if btn_UnlockMatters2 is not None:
    btn_UnlockMatters2.IsChecked = False


def cbo_GroupBy_SelectionChanged(s, event):
  refresh_ListOfLockedMatters()

def btn_RefreshLockedMatters_Click(s, event):
  refresh_ListOfLockedMatters()



class mGroupBy(object):
  def __init__(self, fName, cName):
    self.FriendlyName = fName
    self.CodeName = cName
    return
  
  def __getitem__(self, index):
    if index == 'FName':
      return self.FriendlyName
    elif index == 'CName':
      return self.CodeName
    
def populate_GroupByCombo(s, event):
  xGB = []
  xGB.append(mGroupBy('OurRef', 'EntityRefMN'))
  xGB.append(mGroupBy('Fee Earner', 'FeeEarner'))
  xGB.append(mGroupBy('Department', 'Department'))
  xGB.append(mGroupBy('MRA Name', 'MRA_Name'))
  xGB.append(mGroupBy('(no grouping)', ''))
  cbo_GroupBy.ItemsSource = xGB
  return
  

## Class for Locked Matters DataGrid
class MatterLocks(object):
  def __init__(self, myOurRef, myCLName, myMatDesc, myFE, myDept, myMO, myDSMO, myEntRef, myMatNo, myMRAID, myMRAName, myMRAExpiry, myTTExpDays, myFEemail, myFEforename, myRiskLevel):
    self.EntityRefMN = myOurRef
    self.ClName = myCLName
    self.MatDesc = myMatDesc
    self.FeeEarner = myFE
    self.Department = myDept
    self.MatterOpened = myMO
    self.DaysSinceMatterOpen = myDSMO
    self.EntityRef = myEntRef
    self.MatterNo = myMatNo
    #self.MRA_ExpiryDate = myExpDate
    self.MRA_ID = myMRAID
    self.MRA_Name = myMRAName
    self.MRA_ExpiryCurrent = myMRAExpiry
    self.MRA_TTExpiryDays = myTTExpDays
    self.FEemail = myFEemail
    self.FEforename = myFEforename
    self.matterRiskLevel = myRiskLevel
    return
    
  def __getitem__(self, index):
    if index == 'OurRef':
      return self.EntityRefMN
    elif index == 'ClientName':
      return self.ClName
    elif index == 'MatterDesc':
      return self.MatDesc
    elif index == 'FE':
      return self.FeeEarner
    elif index == 'Dept':
      return self.Department
    elif index == 'MatterOpened':
      return self.MatterOpened     
    elif index == 'DaysSinceOpen':
      return self.DaysSinceMatterOpen
    elif index == 'EntRef':
      return self.EntityRef
    elif index == 'MatNo':
      return self.MatterNo
    #elif index == 'ExpDate':
    #  return self.MRA_ExpiryDate
    elif index == 'MRAID':
      return self.MRA_ID
    elif index == 'MRA Name':
      return self.MRA_Name
    elif index == 'MRA Exp':
      return self.MRA_ExpiryCurrent
    elif index == 'TT ExpDays':
      return self.MRA_TTExpiryDays
    elif index == 'FE_Email':
      return self.FEemail
    elif index == 'FE_Forename':
      return self.FEforename
    elif index == 'MatterRiskLevel':
      return self.matterRiskLevel
    else:
      return ''
      
def refresh_ListOfLockedMatters():

  # first need to get the lock ID for 'LockedByRiskDepartment' lock
  lockID_SQL = """SELECT CASE WHEN EXISTS(SELECT Code FROM Locks WHERE Description = 'LockedByRiskDept') THEN 
                  (SELECT Code FROM Locks WHERE Description = 'LockedByRiskDept') ELSE 0 END """
  lockID = runSQL(lockID_SQL, False, '', '')
  
  if int(lockID) == 0:
    MessageBox.Show("There doesn't appear to be a Lock setup in the name of 'LockedByRiskDept', so cannot list matters locked with this lock!", "Error: Refresh List of Locked Matters...")
    return
  
  # This function will populate the list of locked matters (first main tab)
  mySQL = """SELECT '00-OurRef' = CONCAT(E.ShortCode, '/', CONVERT(nvarchar, M.Number)), 
                    '01-ClientName' = E.LegalName, 
                    '02-MatterDescription' = M.Description, 
                    '03-Fee Earner' = CONCAT('(', U.Code, ') ', U.FullName), 
                    '04-Dept' = CTG.Name, 
                    '05-MatterCreated' = M.Created, 
                    '06-Days since matter open' = DATEDIFF(DAY, M.Created, GETDATE()), 
                    '07-FullLenEntityRef' = M.EntityRef, 
                    '08-MatterNo' = M.Number, 
                    '09-OV_ID' = MRAO.ID, 
                    '10-OV_Name' = MRAO.LocalName, 
                    '11-ExpiryDate' = MRAO.ExpiryDate, 
                    '12-TT ExpiryDays' = TT.ValidityPeriodDays, 
                    '13-FEEmail' = U.EMailExternal, 
                    '14-FE_ForeName' = ISNULL(U.Forename, U.FullName), 
                    '15-MatterRiskLevel' = CASE MRAO.RiskRating WHEN 1 THEN 'Low' WHEN 2 THEN 'Medium' WHEN 3 THEN 'High' ELSE '-unknown-' END 
            FROM Matters M JOIN Entities E ON M.EntityRef = E.Code 
              LEFT OUTER JOIN Users U ON M.FeeEarnerRef = U.Code 
              LEFT OUTER JOIN CaseTypes CT ON M.CaseTypeRef = CT.Code 
              LEFT OUTER JOIN CaseTypeGroups CTG ON CT.CaseTypeGroupRef = CTG.ID 
              LEFT OUTER JOIN EntityMatterLocks EML ON M.EntityRef = EML.EntityRef AND M.Number = EML.MatterNo 
              LEFT OUTER JOIN Usr_MRA_Overview MRAO ON M.EntityRef = MRAO.EntityRef AND M.Number = MRAO.MatterNo AND TypeID IN (SELECT TypeID FROM Usr_MRA_TemplateTypes WHERE Is_MRA = 'Y') 
              LEFT OUTER JOIN Usr_MRA_TemplateTypes TT ON MRAO.TypeID = TT.TypeID 
            WHERE EML.LockID = {0} AND ISNULL(MRAO.Status, '') <> 'Complete' """.format(lockID)   #AND MRAO.ExpiryDate <= GETDATE() "
  
  if cboDept.SelectedIndex > -1:
    mySQL += "AND CTG.Name = '{0}' ".format(cboDept.SelectedItem['Name'])
    
  if cboFeeEarner.SelectedIndex > -1:
    mySQL += "AND U.Code = '{0}' ".format(cboFeeEarner.SelectedItem['Code'])
    
  if len(str(txtEntityRef.Text)) > 4:
    # need to get full length entity ref
    fullLenRef = get_FullEntityRef(str(txtEntityRef.Text))
    mySQL += "AND M.EntityRef = '{0}' ".format(fullLenRef)
    
  if len(str(txtMatterNo.Text)) > 0:
    mySQL += "AND M.Number = {0} ".format(txtMatterNo.Text)
    
  if len(str(txtSearch.Text)) > 0: 
    mySQL += "AND (E.LegalName LIKE '%{0}%' OR M.Description LIKE '%{0}%') ".format(txtSearch.Text)
  
  #MessageBox.Show("SQL: " + mySQL, "Refresh List of Locked Matters...")
  
  _tikitDbAccess.Open(mySQL)
  myItems = []
  
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          iOurRef = '-' if dr.IsDBNull(0) else dr.GetString(0)
          iClName = '-' if dr.IsDBNull(1) else dr.GetString(1)
          iMatDesc = '-' if dr.IsDBNull(2) else dr.GetString(2)
          iFE = '-' if dr.IsDBNull(3) else dr.GetString(3)
          iDept = '-' if dr.IsDBNull(4) else dr.GetString(4)
          iMO = '-' if dr.IsDBNull(5) else dr.GetValue(5)
          iDSMO = 0 if dr.IsDBNull(6) else dr.GetValue(6)
          iEntRef = '-' if dr.IsDBNull(7) else dr.GetString(7)
          iMatNo = '-' if dr.IsDBNull(8) else dr.GetValue(8)
          iMRAid = 0 if dr.IsDBNull(9) else dr.GetValue(9)
          iMRAName = '-' if dr.IsDBNull(10) else dr.GetString(10)
          iMRAExp = 0 if dr.IsDBNull(11) else dr.GetValue(11)
          iTTExpDays = 0 if dr.IsDBNull(12) else dr.GetValue(12)
          iFEemail = '-' if dr.IsDBNull(13) else dr.GetString(13)
          iFEforename = '-' if dr.IsDBNull(14) else dr.GetString(14)
          iRiskLevel = '-' if dr.IsDBNull(15) else dr.GetString(15)

          myItems.append(MatterLocks(iOurRef, iClName, iMatDesc, iFE, iDept, iMO, iDSMO, iEntRef, iMatNo, iMRAid, iMRAName, iMRAExp, iTTExpDays, iFEemail, iFEforename, iRiskLevel))  

      
    dr.Close()
  _tikitDbAccess.Close()
  
  # if allowing user to set different Groupings follow this lead
  if cbo_GroupBy.SelectedIndex == -1:
    gbOption = ''
  else:
    gbOption = cbo_GroupBy.SelectedItem['CName']
    
  if len(gbOption) == 0:
    dg_LockedMatters.ItemsSource = myItems
  else:
    tmpC = ListCollectionView(myItems)
    tmpC.GroupDescriptions.Add(PropertyGroupDescription(gbOption))
    dg_LockedMatters.ItemsSource = tmpC 
  
  if dg_LockedMatters.Items.Count == 0:
    dg_LockedMatters.Visibility = Visibility.Hidden
    tb_NoLockedMatters.Visibility = Visibility.Visible
  else:
    dg_LockedMatters.Visibility = Visibility.Visible
    tb_NoLockedMatters.Visibility = Visibility.Hidden
  return
      
      
class twoColList(object):
  def __init__(self, myCode, myName):
    self.Code = myCode
    self.Name = myName
    return
    
  def __getitem__(self, index):
    if index == 'Code':
      return self.Code
    elif index == 'Name':
      return self.Name
    else:
      return ''
      
def populate_FeeEarnersList(s, event): 
  # This function populates the Fee Earner drop-down list on the 'filter' area of the 'Locked Matters' (Manage Locks) page
  mySQL = "SELECT Code, FullName FROM Users WHERE FeeEarner = 1 AND Locked = 0 AND UserStatus = 0 ORDER BY FullName"
  
  _tikitDbAccess.Open(mySQL)
  myFEitems = []
  
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          myCode = '-' if dr.IsDBNull(0) else dr.GetString(0)
          myName = '-' if dr.IsDBNull(1) else dr.GetString(1)
          myFEitems.append(twoColList(myCode, myName))  
    else:
      myFEitems.append(twoColList('-', '-'))
      
    dr.Close()
  _tikitDbAccess.Close()
  
  cboFeeEarner.ItemsSource = myFEitems
  return
    

def populate_DepartmentList(s, event):
  # This function populates the Department drop-down list on the 'filter' area of the 'Locked Matters' (Manage Locks) page
  dpSQL = "SELECT Name FROM CaseTypeGroups ORDER BY Name"
  
  _tikitDbAccess.Open(dpSQL)
  dpItem = ['Global']
  
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          iDept = '-' if dr.IsDBNull(0) else dr.GetString(0) 
          dpItem.append(iDept)
      
    dr.Close()
  _tikitDbAccess.Close()
  
  cboDept.ItemsSource = dpItem
  return      
      

def clear_LockedMatters_Filters(s, event):
  # This function will clear all the input fields in the 'filter criteria' area and repopulate Locked Matters list
  cboDept.SelectedIndex = -1
  cboFeeEarner.SelectedIndex = -1
  txtSearch.Text = ''
  txtEntityRef.Text = ''
  txtMatterNo.Text = ''
  
  refresh_ListOfLockedMatters()
  return
  

def unlockMatter(withReason = ''):
  # This function will unlock the selected matter(s)
  # AND (as of 29/02/2024) will also extend the 'Expiry Date' for all MRA items by their default time
  # Note, still needs updating to allow selection of many items (button implies we can select many rows and currently do allow multi-select on DG (was 'Extended' now set to 'Single')

  #! MP 10/03/2026: this has been made a little more efficient as we now grab extra data up-front and store against the datagrid item.
  #! (eg: to reduce on the number of separate SQL calls)
  #! Extra fields added: EmailTo (address), emailToUserName, RiskRating (to show as 'Low|Med|High')

  if dg_LockedMatters.SelectedIndex == -1:
    MessageBox.Show("No matter has been selected!\nPlease select a matter before clicking 'Unlock'", "Error: Unlock Selected Matter...")
    return

  if withReason == '':
    MessageBox.Show("No reason for the matter locking has been selected")
    return
  
  selItem = dg_LockedMatters.SelectedItem
  # get entityRef and MatterNo and form SQL to run the stored procedure
  tmpOurRef = selItem['OurRef']
  tmpEntity = selItem['EntRef']
  tmpMatter = selItem['MatNo']
  tmpOVID = selItem['MRAID']
  tmpEmailTo = selItem['FE_Email']
  tmpEmailCC = ''
  tmpToUserName = selItem['FE_Forename']
  tmpMatDesc = selItem['MatterDesc']
  tmpClName = selItem['ClientName']
  tmpAddtl1 = selItem['MRA Name']
  tmpAddtl2 = selItem['RiskRating']
  tmpExpDays = selItem['TT ExpDays']

  unlockCode = "[SQL: EXEC TW_LockHandler '{0}', {1}, 'LockedByRiskDept', 'UnLock']".format(tmpEntity, tmpMatter)
    
  update_log_SQL = "INSERT INTO Usr_Unlock_Log (MatterNo, EntityRef, Reason, Date_Unlocked) VALUES ({0}, '{1}', '{2}', GETDATE())".format(tmpMatter, tmpEntity, withReason.replace("'", "''"))
  runSQL(update_log_SQL, True, "There was an error updating the unlock log", "Error: Unlock Selected Matter - Updating Log...")

  try:
    _tikitResolver.Resolve(unlockCode)
  except:
    MessageBox.Show("There was an error unlocking matter: " + str(tmpOurRef), "Error: Unlock Selected Matter...")
    return
  

  # iterate over items in list and extend ALL MRA's expiry date
  #for dgRow in dg_LockedMatters.Items:
  #  if dgRow.EntityRef == tmpEntity and str(dgRow.MatterNo) == str(tmpMatter):
  #    tmpMRAid = dgRow.MRA_ID
  #    tmpTTDefDays = dgRow.MRA_TTExpiryDays
  #
  #    tmpSQL = "UPDATE Usr_MRA_Overview SET ExpiryDate = DATEADD(day, " + str(tmpTTDefDays) + ", GETDATE()) WHERE ID = " + str(tmpMRAid) 
  #    #MessageBox.Show(tmpSQL, "DEBUGGING")
  #    runSQL(tmpSQL, True, "There was an error extending the Matter Risk Assessment 'Expiry Date'", "Error: Unlock Selected Matter - Extending MRA ExpiryDate...")

  # extend expiry date by number of expiry days (disabled above as unsure why I thought we should extend expiry date for ALL MRAs)
  tmpSQL = "UPDATE Usr_MRA_Overview SET ExpiryDate = DATEADD(day, {0}, GETDATE()) WHERE ID = {1}".format(tmpExpDays, tmpOVID)
  runSQL(tmpSQL, True, "There was an error extending the Matter Risk Assessment 'Expiry Date'", "Error: Unlock Selected Matter - Extending MRA ExpiryDate...")


  # WE ALSO NEED TO TRIGGER THE TASK CENTRE TASK TO NOTIFY THE FE THAT MATTER HAS BEEN UNLOCKED
  tc_Trigger = """INSERT INTO Usr_MRA_Events (Date, UserRef, ActionTrigger, OV_ID, EmailTo, EmailCC, ToUserName, 
                        OurRef, MatterDesc, ClientName, Addtl1, Addtl2, FullEntityRef, Matter_No) 
                VALUES(GETDATE(), '{userRef}', 'Matter_Unlocked', {ovID}, '{emailTo}', '{emailCC}', 
                '{toUserName}', '{ourRef}', '{matDesc}', '{clientName}', '{addtl1}', '{addtl2}', 
                '{entRef}', {matNo})""".format(userRef=_tikitUser, ovID=tmpOVID, 
                                               emailTo=tmpEmailTo, emailCC=tmpEmailCC, toUserName=tmpToUserName, 
                                               ourRef=tmpOurRef, matDesc=tmpMatDesc, clientName=tmpClName, 
                                               addtl1=tmpAddtl1, addtl2=tmpAddtl2, 
                                               entRef=tmpEntity, matNo=tmpMatter)
  runSQL(tc_Trigger, True, "There was an error triggering the Task Center task to notify the Fee Earner that the matter is unlocked", "Error: Task Centre 'Unlocked' notification...")


  MessageBox.Show("Successfully unlocked matter {0} - list will now refresh".format(tmpOurRef), "Unlock Matter - Success...")
  refresh_ListOfLockedMatters()
  return
  

# # # #   END OF:   L O C K E D   M A T T E R S   # # # # 


# # # #   C A S E   T Y P E   D E F A U L T S   # # # #
def add_Missing_CaseTypeDefaults(forWhat=''):
  #! New Added 04/09/2025 - whenever a refresh is called on the 'CaseTypeDefaults' datagrid, we call this function to add any missing
  #!                        CaseTypes to our respective 'CaseTypeDefaults' list (eg: any that have been added since last 'setup')
  #! NB: this is coded to EXCLUDE any CaseType with 'Project' in its name!!

  insertSQL = """INSERT INTO Usr_MRA_CaseType_Defaults (CaseTypeID, TemplateID, TypeName)
                 SELECT CT.Code, -1, '{myFor}'
                  FROM CaseTypes CT
                    LEFT OUTER JOIN Usr_MRA_CaseType_Defaults mCTD ON CT.Code = mCTD.CaseTypeID AND mCTD.TypeName = '{myFor}'
                 WHERE CT.Description NOT LIKE '%Project%' AND ISNULL(mCTD.ID, 0) = 0""".format(myFor=forWhat)

  runSQL(codeToRun=insertSQL, showError=True, 
         errorMsgText="There was an error adding missing CaseTypes for '{myFor}'.\nUsing SQL:\n{mySQL}".format(myFor=forWhat, mySQL=insertSQL), 
         errorMsgTitle="Error: Adding missing CaseTypes...")

  return


class CopyAnswersFrom(object):
  def __init__(self, myCode, myName):
    self.QID = myCode
    self.Name = myName
    return
    
  def __getitem__(self, index):
    if index == 'ID':
      return self.QID
    elif index == 'Name':
      return self.Name
    else:
      return ''



# # # #   M A N A G E   D R O P   D O W N   L I S T S   # # # #

def addNewList(s, event):
  #! Manage Global Answers - Main List - Add New button
  # This function will add a dummy answer with 'newGroup' name and refresh lists
  # Get the next NEGATIVE number (negative numbers are representing our 'global' lists)
  newID = _tikitResolver.Resolve("[SQL: SELECT MIN(QuestionID) - 1 FROM Usr_MRA_TemplateAs]")
  insert_SQL = """INSERT INTO Usr_MRA_TemplateAs (GroupName, AnswerText, Score, DisplayOrder, QuestionID) 
                  VALUES ('NewList', 'example', 0, 1, {0})""".format(newID)
  runSQL(insert_SQL, True, "There was an error adding a new list group", "Error: Add new Answer List group...")
  
  # refresh list and select FIRST item
  refresh_AnswerListGroups(s, event)
  dg_Lists.SelectedIndex = 0  #(dg_Lists.Items.Count - 1)
  return


def duplicateSelectedList(s, event):
  #! Manage Global Answers - Main List - Duplicate Selected button
  # This function will duplicate the currently selected list, as a new list with '(copy)' appended to the name
  # and with a new GLOBAL ID (QuestionID) assigned to it. (NB: copies all current answers)
  if dg_Lists.SelectedIndex == -1:
    MessageBox.Show("Nothing is selected that can be duplicated!", "Error: Duplicate selected list...")
    return

  currID = dg_Lists.SelectedItem['ID']
  newID = _tikitResolver.Resolve("[SQL: SELECT MIN(QuestionID) - 1 FROM Usr_MRA_TemplateAs]")
  insert_SQL = """INSERT INTO Usr_MRA_TemplateAs (GroupName, AnswerText, Score, DisplayOrder, QuestionID) 
                  SELECT GroupName + ' (copy)', AnswerText, Score, DisplayOrder, {0} FROM Usr_MRA_TemplateAs WHERE QuestionID = {1}""".format(newID, currID)
  runSQL(insert_SQL, True, "There was an error duplicating the selected list group", "Error: Duplicate Selected Answer List group...")
  
  # refresh list and select FIRST item
  refresh_AnswerListGroups(s, event)
  dg_Lists.SelectedIndex = 0  #(dg_Lists.Items.Count - 1)
  return


def deleteSelectedList(s, event):
  #! Function being disabled for now - as we shouln't really be able to delete a list that is used in a Question
  #! This is for historic purposes only, because when looking at old data, and if you delete a list/question/answer, it will break the data integrity (no information anymore!!)
  #! Therefore need to be clever with a new solution - firstly, perhaps just have the ability for a MRA/Question/Answer to be able to be 'de-activated' to the effect that it will
  #! not be seen anymore (unless you select 'Show Inactive' or something like that on NMRA Setup perhaps?). This will achive the desired result of Amy not having to see item anymore, 
  #! but we also keep any old data that could be referring to it. We can still have functionality to delete UNUSED items though, but will need checks.
  #! Perhaps then, when Amy wants to amend an NMRA, instead of working directly on the 'Live' one, she makes a duplicate of the one she wishes to edit, and proceed to edit that one.
  #! This also means any historical NMRA's aren't inadvertantly changed by her changes. Maybe we even add 'EffecitiveFrom' date to the NMRA template, to allow Risk to specify when
  #! a new one applies for a Department/CaseType - eg: meaning that she doesn't need to manually go into the 'Case Type Defaults' manually setting the new 'template' to use.
  #! (could also have a button to apply that too).  Also: just thought of, wouldn't it make sense to move the 'CaseType Defaults' to within the editing of the NMRA template,
  #! again, feels like it would flow better - but still keep other existing screen around as this provides an overall of all for the department/case type.
  #! This will be a big change, so will need to think about it more, but I think this is the way to go.

  # This function will delete the currently selected list
  if dg_Lists.SelectedIndex == -1:
    MessageBox.Show("Nothing is selected that can be deleted!", "Error: Delete selected list...")
    return
  
  sel_Name = dg_Lists.SelectedItem['Name']
  sel_ID = dg_Lists.SelectedItem['ID']
  msg = "Are you sure you want to delete the following item:\n{0}?".format(sel_Name)
  myResult = MessageBox.Show(msg, 'Delete item...', MessageBoxButtons.YesNo)
  
  if myResult == DialogResult.Yes:
    # Form the SQL to delete row and execute the SQL 
    Delete_SQL = "[SQL: DELETE FROM Usr_MRA_TemplateAs WHERE QuestionID = '{0}']".format(sel_ID)
    try:
      _tikitResolver.Resolve(Delete_SQL)
      refresh_AnswerListGroups(s, event)
    except:
      MessageBox.Show("There was an error trying to delete answer group, using SQL:\n" + sql_MoveUp, "Error: Delete Selected...")
  return


def ListGroup_SelectionChanged(s, event):
  # This function will update list of 'items'
  if dg_Lists.SelectedIndex > -1:
    lbl_SelectedList.Content = dg_Lists.SelectedItem['Name']
    lbl_SelectedListQID.Content = dg_Lists.SelectedItem['ID']
    refresh_AnswerListItems(s, event)
  else:
    lbl_SelectedList.Content = ''
    lbl_SelectedListQID.Content = ''
  return


def ListGroup_CellEditEnding(s, event):
  # This function will commit changes to the list name (update all records that use this name)
  tmpCol = event.Column
  tmpColName = tmpCol.Header
  
  # Get Initial values updated
  qID = dg_Lists.SelectedItem['ID']
  newName = dg_Lists.SelectedItem['Name']
  newName1 = newName.replace("'", "''")
  oldName = lbl_SelectedList.Content

  # Conditionally add parts depending on column updated and whether value has changed
  if tmpColName != 'List Name':
    return

  if newName != oldName:
    # Added: 25/07/2025 - I've accidentally changed the name of a list without indending to, and when speaking to Amy, she
    # mentioned the same, so I suggested adding a 'confirmation' prompt before blindly changing
    tmpMessage = "Are you sure you want to change the name of the list from:\n{0}\nto:\n{1}?".format(oldName, newName)
    myResult = MessageBox.Show(tmpMessage, "Change List Name", MessageBoxButtons.YesNo)
    if myResult != DialogResult.Yes:
      updateSQL = "UPDATE Usr_MRA_TemplateAs SET GroupName = '" + str(newName1) + "' WHERE QuestionID = " + str(qID) 
      runSQL(updateSQL, True, "There was an error trying to update Answer list item name", "Error: Updating Answer list item...")

    # we refresh the list of items to make sure we show what the actual name is (saved new name / retained old name)
    refresh_AnswerListItems(s, event)
  return


def ListItems_SelectionChanged(s, event):
  # This function temp stores selected item
  if dg_ListItems.SelectedIndex > -1:
    lbl_ListItemText.Content = dg_ListItems.SelectedItem['Item Text']
    lbl_ListItemScore.Content = dg_ListItems.SelectedItem['Score']
    lbl_ListItemEC.Content = dg_ListItems.SelectedItem['EmailComment']
  else:
    lbl_ListItemText.Content = ''
    lbl_ListItemScore.Content = ''
    lbl_ListItemEC.Content = ''
  return

  
def ListItems_CellEditEnding(s, event):
  # This function commits changes to list item
  tmpCol = event.Column
  tmpColName = tmpCol.Header
  
  # Get Initial values updated
  updateSQL = '[SQL: UPDATE Usr_MRA_TemplateAs SET '
  countOfUpdates = 0
  itemID = dg_ListItems.SelectedItem['RowID']
  newName = dg_ListItems.SelectedItem['Item Text']
  newName1 = newName.replace("'", "''")
  newScore = dg_ListItems.SelectedItem['Score']
  newEC = dg_ListItems.SelectedItem['EmailComment']

  if itemID == 'x':
    return

  # Conditionally add parts depending on column updated and whether value has changed
  if tmpColName == 'Item Text':
    if newName != lbl_ListItemText.Content:
      updateSQL += "AnswerText = '{0}' ".format(newName1)
      countOfUpdates += 1

  elif tmpColName == 'Item Score':
    if newScore != lbl_ListItemScore.Content:
      updateSQL += "Score = {0} ".format(newScore)
      countOfUpdates += 1

  elif tmpColName == 'Email Comment':
    if newEC != lbl_ListItemEC.Content:
      updateSQL += "EmailComment = '{0}' ".format(newEC.replace("'","''"))
      countOfUpdates += 1

  # Add WHERE clause
  updateSQL += "WHERE ID = {0}]".format(itemID)
  #MessageBox.Show("UpdateSQL = " + updateSQL + "\nNewEC (list item): " + str(newEC) + "\nPrev value (in label): " + str(lbl_ListItemEC.Content))

  # Only run if something was changed
  if countOfUpdates > 0:
    runSQL(updateSQL, True, "There was an error trying to update Answer list item", "Error: Updating Answer list item...")
    refresh_AnswerListItems(s, event)
  return


def refresh_AnswerListGroups(s, event):
  # This function refreshes the 'ANSWER LIST GROUP' data drid
  # Currently doing tripple-duty as updates: DataGrid (Manage Answers tab), 'Copy Answers From' on Editing Qs MRA, and 'Answer List' on Editing Qs on FR
  # 'Manage Answers > Global Answers' is now only for editing 'Global' (or template) Answers and shouldn't show items for existing Q's
  # No longer interested in 'Count Used' as these are only TEMPLATE answers to be COPIED onto actual Questions
  
  getTableSQL = """SELECT CASE WHEN TAs.QuestionID < 0 THEN TAs.GroupName ELSE (SELECT QuestionText FROM Usr_MRA_TemplateQs WHERE ID = TAs.QuestionID) END, 
                  TAs.QuestionID, CASE WHEN TAs.QuestionID < 0 THEN '(GLOBAL) ' + TAs.GroupName ELSE (SELECT QuestionText FROM Usr_MRA_TemplateQs WHERE ID = TAs.QuestionID) END 
                  FROM Usr_MRA_TemplateAs TAs GROUP BY GroupName, QuestionID ORDER BY QuestionID"""
  
  tmpItem = []
  tmpItem2 = []
  
  _tikitDbAccess.Open(getTableSQL)
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          tmpGName = '' if dr.IsDBNull(0) else dr.GetString(0)
          tmpID = 0 if dr.IsDBNull(1) else dr.GetValue(1)
          tmpGName2 = '' if dr.IsDBNull(2) else dr.GetString(2)

          # new - as we only want to see 'Global' template Answers on Manage Answers tab, we'll only add to list if ID is below zero
          if tmpID < 0:
            tmpItem.append(CopyAnswersFrom(tmpID, tmpGName))
          # we'll still want to see all other questions (I think) in 'Copy Answers From' drop-downs, so always add these
          tmpItem2.append(CopyAnswersFrom(tmpID, tmpGName2))

    dr.Close()
  
  #close db connection
  _tikitDbAccess.Close()
  dg_Lists.ItemsSource = tmpItem
  #cbo_QuestionAnswerList.ItemsSource = tmpItem2
  
  if dg_Lists.Items.Count == 0:
    lbl_NoGlobalGroups.Visibility = Visibility.Visible
    dg_Lists.Visibility = Visibility.Hidden
  else:
    lbl_NoGlobalGroups.Visibility = Visibility.Collapsed
    dg_Lists.Visibility = Visibility.Visible
  return 


class AnswerListItems(object):
  def __init__(self, myDO, myText, myScore, myRowID, myEC):
    self.liDO = myDO
    self.liText = myText
    self.liScore = myScore
    self.liCode = myRowID
    self.liEmailComment = myEC
    return
    
  def __getitem__(self, index):
    if index == 'Order':
      return self.liDO
    elif index == 'Item Text':
      return self.liText
    elif index == 'Score':
      return self.liScore
    elif index == 'RowID':
      return self.liCode
    elif index == 'EmailComment':
      return self.liEmailComment
    else: 
      return ''


def refresh_AnswerListItems(s, event):
  # This function refreshes the 'ANSWER LIST ITEMS' data grid
  selID = dg_Lists.SelectedItem['ID']
  getTableSQL = "SELECT DisplayOrder, AnswerText, Score, ID, EmailComment FROM Usr_MRA_TemplateAs WHERE QuestionID = {0} ORDER BY DisplayOrder".format(selID)
  
  tmpItem = []
  _tikitDbAccess.Open(getTableSQL)
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          tmpDO = 0 if dr.IsDBNull(0) else dr.GetValue(0)
          tmpAText = '' if dr.IsDBNull(1) else dr.GetString(1)
          tmpScore = 0 if dr.IsDBNull(2) else dr.GetValue(2)
          tmpID = 0 if dr.IsDBNull(3) else dr.GetValue(3)
          tmpEC = '' if dr.IsDBNull(4) else dr.GetString(4)

          tmpItem.append(AnswerListItems(tmpDO, tmpAText, tmpScore, tmpID, tmpEC))
    else:
      tmpItem.append(AnswerListItems(0, 'No items currently exist', 0, 0, ''))

    dr.Close()
  
  #close db connection
  _tikitDbAccess.Close()
  dg_ListItems.ItemsSource = tmpItem
  set_Visibility_ofAnswerItemsDG()
  return


def set_Visibility_ofAnswerItemsDG():
  # This function will hide the Answer list items datagrid if no items exist and show a help label
  
  if dg_ListItems.Items.Count == 0:
    dg_ListItems.Visibility = Visibility.Hidden
    tb_NoAnswerItems.Visibility = Visibility.Visible
  else:
    dg_ListItems.Visibility = Visibility.Visible
    tb_NoAnswerItems.Visibility = Visibility.Hidden
  return


def addNewListItem(s, event):
  #! [Manage Global Answers] tab > [Global Answers...] tab > [2. Edit List Items] section (DataGrid on right... the items in the main 'global' list)
  # This function will add a new list item to the currently selected group
  selID = dg_Lists.SelectedItem['ID']
  newDO = _tikitResolver.Resolve("[SQL: SELECT ISNULL(MAX(DisplayOrder) + 1, 1) FROM Usr_MRA_TemplateAs WHERE QuestionID = {0}]".format(selID))
  insert_SQL = """INSERT INTO Usr_MRA_TemplateAs (GroupName, AnswerText, Score, DisplayOrder, QuestionID) 
                  VALUES('{0}', '(new)', 0, {1}, {2})""".format(lbl_SelectedList.Content, newDO, selID)
  runSQL(insert_SQL, True, "There was an error adding a new list item", "Error: Add New Answer List item...")
  
  # auto select last item...
  refresh_AnswerListItems(s, event)
  dg_ListItems.SelectedIndex = (dg_ListItems.Items.Count - 1)
  return
  
  
def duplicateSelectedListItem(s, event):
  #! [Manage Global Answers] tab > [Global Answers...] tab > [2. Edit List Items] section (DataGrid on right... the items in the main 'global' list)
  # This function will duplicate the currently selected list item
  if dg_ListItems.SelectedIndex == -1:
    MessageBox.Show("Nothing selected to duplicate!", "Error: Duplicating Selected Answer item...")
    return
  
  selectedID = dg_ListItems.SelectedItem['ID']
  insert_SQL = """INSERT INTO Usr_MRA_TemplateAs (GroupName, AnswerText, Score, DisplayOrder, EmailComment, QuestionID) 
                  SELECT TA.GroupName, TA.AnswerText + ' (copy)', TA.Score, (SELECT MAX(TA1.DisplayOrder) + 1 FROM Usr_MRA_TemplateAs TA1 WHERE TA1.GroupName = TA.GroupName), 
                    EmailComment, QuestionID FROM Usr_MRA_TemplateAs TA WHERE TA.ID = {0}""".format(selectedID)   
  runSQL(insert_SQL, True, "There was an error duplicating the selected list item", "Error: Duplicate Selected Answer List item...")
  
  # auto select last item...
  refresh_AnswerListItems(s, event)
  dg_ListItems.SelectedIndex = (dg_ListItems.Items.Count - 1)
  return

  
def deleteSelectedListItem(s, event):
  #! [Manage Global Answers] tab > [Global Answers...] tab > [2. Edit List Items] section (DataGrid on right... the items in the main 'global' list)
  
  #! As mentioned on other 'Delete' functions - we really need to re-consider this functionality as it will break data integrity
  #! We should really be able to 'de-activate' items, rather than delete, and hide from view (unless 'Show Inactive' is ticked)
  #! (this prevents ruining old completed NMRAs). So two things, get/display count used so that we can check that and prevent delete until
  #! 2) we provide option to update all 'old' NMRAs to a new value or something)

  # This function will delete the currently selected list item
  tmpNewFocusRow = dgItem_DeleteSelected(dgControl=dg_ListItems, 
                                         tableToUpdate='Usr_MRA_TemplateAs', 
                                         sqlOrderColName='DisplayOrder', 
                                         dgIDcolName='RowID', 
                                         dgOrderColName='Order', 
                                         dgNameDescColName='Item Text', 
                                         sqlOtherCheckCol='QuestionID', 
                                         sqlOtherCheckValue=lbl_SelectedListQID.Content)
  if tmpNewFocusRow > -1:
    refresh_AnswerListItems(s, event)
    dg_ListItems.Focus()
    dg_ListItems.SelectedIndex = tmpNewFocusRow  
  return
  
  
def Answers_MoveItemToTop(s, event):
  #! [Manage Global Answers] tab > [Global Answers...] tab > [2. Edit List Items] section (DataGrid on right... the items in the main 'global' list)
  # This function will move the selected ANSWER to the top row (and all other items down one)
  tmpNewFocusRow = dgItem_MoveToTop(dgControl=dg_ListItems, 
                                    tableToUpdate='Usr_MRA_TemplateAs', 
                                    sqlOrderColName='DisplayOrder', 
                                    dgIDcolName='RowID', 
                                    dgOrderColName='Order', 
                                    sqlOtherCheckCol='QuestionID', 
                                    sqlOtherCheckValue=lbl_SelectedListQID.Content)
  if tmpNewFocusRow > -1:
    refresh_AnswerListItems(s, event)
    dg_ListItems.Focus()
    dg_ListItems.SelectedIndex = tmpNewFocusRow
  return
  
  
def Answers_MoveItemUp(s, event):
  #! [Manage Global Answers] tab > [Global Answers...] tab > [2. Edit List Items] section (DataGrid on right... the items in the main 'global' list)
  # This function will move the selected ANSWER up one row (and all other items down one)
  tmpNewFocusRow = dgItem_MoveUp(dgControl=dg_ListItems, 
                                 tableToUpdate='Usr_MRA_TemplateAs', 
                                 sqlOrderColName='DisplayOrder', 
                                 dgIDcolName='RowID', 
                                 dgOrderColName='Order', 
                                 sqlOtherCheckCol='QuestionID', 
                                 sqlOtherCheckValue=lbl_SelectedListQID.Content)
  if tmpNewFocusRow > -1: 
    refresh_AnswerListItems(s, event)
    dg_ListItems.Focus()
    dg_ListItems.SelectedIndex = tmpNewFocusRow
  return
  

def Answers_MoveItemDown(s, event):
  #! [Manage Global Answers] tab > [Global Answers...] tab > [2. Edit List Items] section (DataGrid on right... the items in the main 'global' list)
  # This function will move the selected Answer down one row (and all other items up one)
  tmpNewFocusRow = dgItem_MoveDown(dgControl=dg_ListItems, 
                                   tableToUpdate='Usr_MRA_TemplateAs', 
                                   sqlOrderColName='DisplayOrder', 
                                   dgIDcolName='RowID', 
                                   dgOrderColName='Order', 
                                   sqlOtherCheckCol='QuestionID', 
                                   sqlOtherCheckValue=lbl_SelectedListQID.Content)
  if tmpNewFocusRow > -1: 
    refresh_AnswerListItems(s, event)
    dg_ListItems.Focus()
    dg_ListItems.SelectedIndex = tmpNewFocusRow  
  return
  
  
def Answers_MoveItemToBottom(s, event):
  #! [Manage Global Answers] tab > [Global Answers...] tab > [2. Edit List Items] section (DataGrid on right... the items in the main 'global' list)
  # This function will move the selected Answer to the bottom row (and all other items up one)
  tmpNewFocusRow = dgItem_MoveToBottom(dgControl=dg_ListItems, 
                                       tableToUpdate='Usr_MRA_TemplateAs', 
                                       sqlOrderColName='DisplayOrder', 
                                       dgIDcolName='RowID', 
                                       dgOrderColName='Order', 
                                       sqlOtherCheckCol='QuestionID', 
                                       sqlOtherCheckValue=lbl_SelectedListQID.Content)
  if tmpNewFocusRow > -1:
    refresh_AnswerListItems(s, event)
    dg_ListItems.Focus()
    dg_ListItems.SelectedIndex = tmpNewFocusRow  
  return
  


class SectionGroups(object):
  def __init__(self, myDO, myDesc, myID, myGroup):
    self.gDO = myDO
    self.gDesc = myDesc
    self.gGroup = myGroup
    self.gID = myID
    return
    
  def __getitem__(self, index):
    if index == 'Order':
      return self.gDO
    elif index == 'Desc':
      return self.gDesc
    elif index == 'Group':
      return self.gGroup
    elif index == 'ID':
      return self.gID
    else: 
      return ''


def refresh_GroupItems(s, event):
  #! [Manage Global Answers] tab > [Groups] tab 
  # ! Linked to XAML control: dg_GroupItems  (Manage Global Answers > Groups tab)
  # This function refreshes the 'Group Items' data grid
  getTableSQL = "SELECT DisplayOrder, ID, Name, ForWhat FROM Usr_MRA_QGroups ORDER BY ForWhat, DisplayOrder"
  
  tmpItem = []

  _tikitDbAccess.Open(getTableSQL)
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          tmpDO = 0 if dr.IsDBNull(0) else dr.GetValue(0)
          tmpID = 0 if dr.IsDBNull(1) else dr.GetValue(1)
          tmpDesc = '' if dr.IsDBNull(2) else dr.GetString(2)
          tmpGroup = '' if dr.IsDBNull(3) else dr.GetString(3)

          tmpItem.append(SectionGroups(tmpDO, tmpDesc, tmpID, tmpGroup))
    dr.Close()
  #close db connection
  _tikitDbAccess.Close()

  # now we have all our items in a propert Python list, we can add to DataGrid (and add 'Grouping')
  tmpC = ListCollectionView(tmpItem)
  tmpC.GroupDescriptions.Add(PropertyGroupDescription("gGroup"))
  dg_GroupItems.ItemsSource = tmpC #tmpItem
  
  if dg_GroupItems.Items.Count == 0:
    dg_GroupItems.Visibility = Visibility.Hidden
    tb_NoGroupItems.Visibility = Visibility.Visible
  else:
    dg_GroupItems.Visibility = Visibility.Visible
    tb_NoGroupItems.Visibility = Visibility.Hidden

  # finally update the drop-downs on the 'Edit Questions' tabs
  populate_FR_QGroups(s, event)
  return



def Group_SelectionChanged(s, event):
  #! [Manage Global Answers] tab > [Groups] tab 
  # This function temp stores selected item
  if dg_GroupItems.SelectedIndex > -1:
    lbl_GroupItemText.Content = dg_GroupItems.SelectedItem['Desc']
    lbl_GroupItemGroup.Content = dg_GroupItems.SelectedItem['Group']
  else:
    lbl_GroupItemText.Content = ''
    lbl_GroupItemGroup.Content = ''
  return


def Group_CellEditEnding(s, event):
  #! [Manage Global Answers] tab > [Groups] tab 
  # This function commits changes to list item
  tmpCol = event.Column
  tmpColName = tmpCol.Header
  
  # Get Initial values updated
  updateSQL = "[SQL: UPDATE Usr_MRA_QGroups SET "
  countOfUpdates = 0
  itemID = dg_GroupItems.SelectedItem['ID']
  newName = dg_GroupItems.SelectedItem['Desc']
  newName1 = newName.replace("'", "''")
  newGroup = dg_GroupItems.SelectedItem['Group']
  newGroup = newGroup.replace("'", "''")

  if itemID != 'x':
    # Conditionally add parts depending on column updated and whether value has changed
    if tmpColName == 'Description':
      if newName != lbl_GroupItemText.Content:
        updateSQL += "Name = '{0}' ".format(newName1)
        countOfUpdates += 1
    if tmpColName == 'Group':
      if newGroup != lbl_GroupItemGroup.Content:
        updateSQL += "ForWhat = '{0}' ".format(newGroup)
        countOfUpdates += 1
    # Add WHERE clause
    updateSQL += "WHERE ID = {0}]".format(itemID)
  
    # Only run if something was changed
    if countOfUpdates > 0:
      try:
        _tikitResolver.Resolve(updateSQL)
      except:
        MessageBox.Show("There was an error trying to update Section/Group item, with SQL:\n" + updateSQL, "Error: Updating 'Section/Group' item...")
      refresh_GroupItems(s, event)
  return


def addNewGroup(s, event):
  #! [Manage Global Answers] tab > [Groups] tab 
  # This function will add a new GROUP item for Matter Risk Assessments
  newDO = _tikitResolver.Resolve("[SQL: SELECT ISNULL(MAX(DisplayOrder) + 1, 1) FROM Usr_MRA_QGroups WHERE ForWhat = 'MRA']")
  insert_SQL = """INSERT INTO Usr_MRA_QGroups (Name, DisplayOrder, ForWhat) 
                  VALUES('(new)', {0}, 'MRA')""".format(newDO)
  runSQL(insert_SQL, True, "There was an error adding a new Group", "Error: Add New Group item...")
  
  # auto select last item...
  refresh_GroupItems(s, event)
  dg_GroupItems.SelectedIndex = (dg_GroupItems.Items.Count - 1)
  return
  
def addNewGroupFR(s, event):
  #! [Manage Global Answers] tab > [Groups] tab 
  # This function will add a new GROUP item for File Reviews
  newDO = _tikitResolver.Resolve("[SQL: SELECT ISNULL(MAX(DisplayOrder) + 1, 1) FROM Usr_MRA_QGroups WHERE ForWhat = 'FR']")
  insert_SQL = """INSERT INTO Usr_MRA_QGroups (Name, DisplayOrder, ForWhat) 
                  VALUES('(new)', {0}, 'FR')""".format(newDO)
  runSQL(insert_SQL, True, "There was an error adding a new Group", "Error: Add New Group item...")
  
  # auto select last item...
  refresh_GroupItems(s, event)
  dg_GroupItems.SelectedIndex = (dg_GroupItems.Items.Count - 1)
  return  
  
def duplicateSelectedGroup(s, event):
  #! [Manage Global Answers] tab > [Groups] tab 
  # This function will duplicate the currently selected list item
  if dg_GroupItems.SelectedIndex == -1:
    MessageBox.Show("Nothing selected to duplicate!", "Error: Duplicating Selected Group...")
    return
  
  selectedID = dg_GroupItems.SelectedItem['ID']
  insert_SQL = """INSERT INTO Usr_MRA_QGroups (Name, DisplayOrder, ForWhat) 
                  SELECT TA.Name + ' (copy)', (SELECT MAX(TA1.DisplayOrder) + 1 FROM Usr_MRA_QGroups TA1 WHERE TA1.ForWhat = TA.ForWhat), TA.ForWhat 
                  FROM Usr_MRA_QGroups TA WHERE TA.ID = {0}""".format(selectedID) 
  runSQL(insert_SQL, True, "There was an error duplicating the selected Group", "Error: Duplicate Selected Group item...")
  
  # auto select last item...
  refresh_GroupItems(s, event)
  dg_GroupItems.SelectedIndex = (dg_GroupItems.Items.Count - 1)
  return

  
def deleteSelectedGroup(s, event):
  #! [Manage Global Answers] tab > [Groups] tab 
  #! NOTE: Another 'Delete' function - see notes added to other 'Delete' ('def deleteSelectedList') for new plan
  # This function will delete the currently selected list item
  tmpGroup = "'{0}'".format(dg_GroupItems.SelectedItem['Group'])
  tmpNewFocusRow = dgItem_DeleteSelected(dgControl=dg_GroupItems, 
                                         tableToUpdate='Usr_MRA_QGroups', 
                                         sqlOrderColName='DisplayOrder', 
                                         dgIDcolName='ID', 
                                         dgOrderColName='Order', 
                                         dgNameDescColName='Desc', 
                                         sqlOtherCheckCol='ForWhat', 
                                         sqlOtherCheckValue=tmpGroup)
  if tmpNewFocusRow > -1:
    refresh_GroupItems(s, event)
    dg_GroupItems.Focus()
    dg_GroupItems.SelectedIndex = tmpNewFocusRow  
  return
  
  
def Group_MoveItemUp(s, event):
  #! [Manage Global Answers] tab > [Groups] tab 
  # This function will move the selected ANSWER up one row (and all other items down one)
  #                dgItem_MoveUp(   dgControl,        tableToUpdate, sqlOrderColName, dgIDcolName, dgOrderColName, sqlOtherCheckCol, sqlOtherCheckValue)
  tmpGroup = "'{0}'".format(dg_GroupItems.SelectedItem['Group'])
  tmpNewFocusRow = dgItem_MoveUp(dgControl=dg_GroupItems, 
                                 tableToUpdate='Usr_MRA_QGroups', 
                                 sqlOrderColName='DisplayOrder',
                                 dgIDcolName='ID', 
                                 dgOrderColName='Order', 
                                 sqlOtherCheckCol='ForWhat', 
                                 sqlOtherCheckValue=tmpGroup)
  if tmpNewFocusRow > -1: 
    refresh_GroupItems(s, event)
    dg_GroupItems.Focus()
    dg_GroupItems.SelectedIndex = tmpNewFocusRow
  return
  

def Group_MoveItemDown(s, event):
  #! [Manage Global Answers] tab > [Groups] tab 
  # This function will move the selected Answer down one row (and all other items up one)
  #                dgItem_MoveDown(   dgControl,        tableToUpdate, sqlOrderColName, dgIDcolName, dgOrderColName, sqlOtherCheckCol, sqlOtherCheckValue)
  tmpGroup = "'{0}'".format(dg_GroupItems.SelectedItem['Group'])
  tmpNewFocusRow = dgItem_MoveDown(dgControl=dg_GroupItems, 
                                   tableToUpdate='Usr_MRA_QGroups', 
                                   sqlOrderColName='DisplayOrder', 
                                   dgIDcolName='ID', 
                                   dgOrderColName='Order', 
                                   sqlOtherCheckCol='ForWhat', 
                                   sqlOtherCheckValue=tmpGroup)
  if tmpNewFocusRow > -1: 
    refresh_GroupItems(s, event)
    dg_GroupItems.Focus()
    dg_GroupItems.SelectedIndex = tmpNewFocusRow  
  return
  
  
# Move to Top / Bottom not practical here as those functions don't take into account the 'grouping' (it gets total number of items in DG, rather than just those in the group)
# Don't want to spend time fixing up now, so will just disable those buttons for now.  
def Group_MoveItemToTop(s, event):
  #! [Manage Global Answers] tab > [Groups] tab 
  # This function will move the selected ANSWER to the top row (and all other items down one)
  #                dgItem_MoveToTop(   dgControl,        tableToUpdate, sqlOrderColName, dgIDcolName, dgOrderColName, sqlOtherCheckCol, sqlOtherCheckValue)
  tmpGroup = "'{0}'".format(dg_GroupItems.SelectedItem['Group'])
  tmpNewFocusRow = dgItem_MoveToTop(dgControl=dg_GroupItems, 
                                    tableToUpdate='Usr_MRA_QGroups', 
                                    sqlOrderColName='DisplayOrder', 
                                    dgIDcolName='ID', 
                                    dgOrderColName='Order', 
                                    sqlOtherCheckCol='ForWhat', 
                                    sqlOtherCheckValue=tmpGroup)
  if tmpNewFocusRow > -1:
    refresh_GroupItems(s, event)
    dg_GroupItems.Focus()
    dg_GroupItems.SelectedIndex = tmpNewFocusRow
  return
  
def Group_MoveItemToBottom(s, event):
  #! [Manage Global Answers] tab > [Groups] tab 
  # This function will move the selected Answer to the bottom row (and all other items up one)
  #                dgItem_MoveToBottom(   dgControl,        tableToUpdate, sqlOrderColName, dgIDcolName, dgOrderColName, sqlOtherCheckCol, sqlOtherCheckValue)
  tmpGroup = "'{0}'".format(dg_GroupItems.SelectedItem['Group'])
  tmpNewFocusRow = dgItem_MoveToBottom(dgControl=dg_GroupItems, 
                                       tableToUpdate='Usr_MRA_QGroups', 
                                       sqlOrderColName='DisplayOrder',
                                       dgIDcolName='ID', 
                                       dgOrderColName='Order', 
                                       sqlOtherCheckCol='ForWhat', 
                                       sqlOtherCheckValue=tmpGroup)
  if tmpNewFocusRow > -1:
    refresh_GroupItems(s, event)
    dg_GroupItems.Focus()
    dg_GroupItems.SelectedIndex = tmpNewFocusRow  
  return



# # # #   END OF:  M A N A G E   D R O P   D O W N   L I S T S   # # # #


# # # #   C O N F I G U R E   F I L E   R E V I E W S   TAB   # # # #

class FR_Templates(object):
  def __init__(self, myCode, myFName, myCountUsed, myQCount):
    self.mraT_Code = myCode
    self.mraT_Desc = myFName
    self.CountUsed = myCountUsed
    self.QCount = myQCount
    return

  def __getitem__(self, index):
    if index == 'Code':
      return self.mraT_Code
    elif index == 'Name':
      return self.mraT_Desc
    elif index == 'CountUsed':
      return self.CountUsed
    elif index == 'QCount':
      return self.QCount
    else:
      return ''


def refresh_FR_Templates(s, event):
  # This funtion populates the main File Review data grid (and also populates the combo drop-downs in the 'Department' and 'Case Type' defaults area)
  # Controls affected with this function: dg_FR_Templates; cbo_FR_Department_TemplateToUse; cbo_FR_CaseType_TemplateToUse
  
  # SQL to populate datagrid
  getTableSQL = """SELECT MRA_TT.TypeID, MRA_TT.TypeName, 'Count Used' = (SELECT COUNT(ID) FROM Usr_MRA_Overview WHERE TypeID = MRA_TT.TypeID), 
                         'QCount' = (SELECT COUNT(ID) FROM Usr_MRA_TemplateQs TQs WHERE TQs.TypeID = MRA_TT.TypeID) 
                   FROM Usr_MRA_TemplateTypes MRA_TT WHERE MRA_TT.Is_MRA = 'N' ORDER BY MRA_TT.ID"""

  tmpItem = []
  tmpItemDDL = []
  tmpItemDDL.append(FR_Templates(-1, '(none)', 0, 0))
  
  _tikitDbAccess.Open(getTableSQL)
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          tmpID = 0 if dr.IsDBNull(0) else dr.GetValue(0)
          tmpName = '' if dr.IsDBNull(1) else dr.GetString(1)
          tmpCU = 0 if dr.IsDBNull(2) else dr.GetValue(2)
          tmpQCount = 0 if dr.IsDBNull(3) else dr.GetValue(3)
          tmpItem.append(FR_Templates(tmpID, tmpName, tmpCU, tmpQCount))
          tmpItemDDL.append(FR_Templates(tmpID, tmpName, tmpCU, tmpQCount))
    else:
      tmpItem.append(FR_Templates(-1, 'No items currently exist - click + to add', 0, 0))

    dr.Close()
  
  #close db connection
  _tikitDbAccess.Close()
  dg_FR_Templates.ItemsSource = tmpItem
  # also populate the drop-down list boxes
  cbo_FR_Department_TemplateToUse.ItemsSource = tmpItemDDL
  cbo_FR_CaseType_TemplateToUse.ItemsSource = tmpItemDDL
  return


def AddNew_FR_Template(s, event):
  # This function will add a new row to the 'File Review' data drid

  #! ADDED: 29/09/2025: get next new TypeID so we can pass directly into INSERT statement
  newTypeID = _tikitResolver.Resolve("[SQL: SELECT ISNULL(MAX(TypeID) + 1, 1) FROM Usr_MRA_TemplateTypes]")
  insertSQL = "[SQL: INSERT INTO Usr_MRA_TemplateTypes (TypeName, Is_MRA, TypeID) VALUES ('File Review (new)', 'N', {0})]".format(newTypeID)
  try:
    _tikitResolver.Resolve(insertSQL)
  except:
    MessageBox.Show("There was an error adding new File Review, using SQL:\n" + str(insertSQL), "Error: Adding new File Review...")
    return
    
  # refresh data grid and select last item
  refresh_FR_Templates(s, event)
  dg_FR_Templates.Focus()
  dg_FR_Templates.SelectedIndex = (dg_FR_Templates.Items.Count - 1)  
  return


def Duplicate_FR_Template(s, event):
  # This function will duplicate the selected File Review (including the questions)
  # Linked to button on XAML: btn_CopySelected_FRTemplate

  if dg_FR_Templates.SelectedIndex == -1:
    MessageBox.Show("Nothing selected to copy!", "Error: Duplicate Selected File Review...")
    return
  
  idItemToCopy = dg_FR_Templates.SelectedItem['Code']
  nameToCopy = dg_FR_Templates.SelectedItem['Name']
  
  # Firstly, copy main template and get new ID
  tempName = "File Review (copy of " + str(idItemToCopy) + ")"
  #! ADDED: 29/09/2025: get next new TypeID so we can pass directly into INSERT statement
  newTypeID = _tikitResolver.Resolve("[SQL: SELECT ISNULL(MAX(TypeID) + 1, 1) FROM Usr_MRA_TemplateTypes]")
  insertSQL = "[SQL: INSERT INTO Usr_MRA_TemplateTypes (TypeName, Is_MRA, TypeID) VALUES ('{0}', 'N', {0})]".format(tempName, newTypeID)
  _tikitResolver.Resolve(insertSQL)
  
  # now get ID of added row...
  #rowID = _tikitResolver.Resolve("[SQL: SELECT TOP 1 ID FROM Usr_MRA_TemplateTypes WHERE TypeName = '{0}']".format(tempName))
  #! ^ do we need this?  Could we not just use the 'newTypeID' from above?
  
  # Then copy over questions (but noting the new ID)
  #if int(rowID) > 0:
  copyQ_SQL = """INSERT INTO Usr_MRA_TemplateQs (DisplayOrder, QuestionText, AnswerList, TypeID) 
                  SELECT DisplayOrder, QuestionText, AnswerList, {0} FROM Usr_MRA_TemplateQs WHERE TypeID = {1}""".format(newTypeID, idItemToCopy)
  #! ^ formally was passing 'rowID', but we shouldn't be using that as we created a 'TypeID' field
  try:
    _tikitResolver.Resolve("[SQL: {0}]".format(copyQ_SQL.strip()))
    MessageBox.Show("Successfully copied '{0}'".format(nameToCopy))
    dg_FR_Templates.Focus()
    dg_FR_Templates.SelectedIndex = (dg_FR_Templates.Items.Count - 1)
  except:
    MessageBox.Show("An error occurred copying the Questions")
  
  refresh_FR_Templates(s, event)
  return  
  
  
def Delete_FR_Template(s, event):
  # This function will delete the selected File Review template (and any questions associated to it)
  # Linked to button on XAML: btn_DeleteSelected_FRTemplate
  
  if dg_FR_Templates.SelectedIndex == -1:
    MessageBox.Show("Nothing selected to delete!", "Error: Delete Selected File Review...")
    return

  # First get the ID, as we'll also want to delete questions using this ID
  tmpID = dg_FR_Templates.SelectedItem['Code'] 

  # Call generic function to do main delete
  tmpNewFocusRow = dgItem_DeleteSelected(dgControl=dg_FR_Templates, 
                                         tableToUpdate='Usr_MRA_TemplateTypes', 
                                         sqlOrderColName='', 
                                         dgIDcolName='ID', 
                                         dgOrderColName='', 
                                         dgNameDescColName='Name', 
                                         sqlOtherCheckCol='', 
                                         sqlOtherCheckValue='')
  if tmpNewFocusRow > -1:
    refresh_FR_Templates(s, event)
    dg_FR_Templates.Focus()
    dg_FR_Templates.SelectedIndex = tmpNewFocusRow
    
    # now to delete all ANSWERS associated to questions with this this ID, then delete the QUESTIONS
    deleteA_SQL = "DELETE FROM Usr_MRA_TemplateAs WHERE QuestionID IN (SELECT ID FROM Usr_MRA_TemplateQs WHERE TypeID = {0})".format(tmpID)
    runSQL(deleteA_SQL, True, "There was an error deleting the Answers associated to the Questions used for the selected File Review", "Error: Deleting File Review Template...")     
    deleteQ_SQL = "DELETE FROM Usr_MRA_TemplateQs WHERE TypeID = {0}".format(tmpID)
    runSQL(deleteQ_SQL, True, "There was an error deleting the Questions associated to the selected File Review", "Error: Deleting File Review Template...")
  return


def Preview_FR_Template(s, event):
  # This function will load the 'Preview' tab (made to look like 'matter-level' XAML) for the selected item
  # Linked to button on XAML: btn_Preview_FRTemplate
  
  if dg_FR_Templates.SelectedIndex == -1:
    MessageBox.Show("Nothing selected to Preview!", "Error: Preview selected File Review...")
    return  
  
  # first need to load up questions onto tab and then display/select tab
  lbl_FR_Preview_ID.Content = dg_FR_Templates.SelectedItem['Code']
  lbl_FR_Preview_Name.Content = dg_FR_Templates.SelectedItem['Name']
  
  # clear existing list
  _tikitResolver.Resolve("[SQL: DELETE FROM Usr_FR_Preview WHERE ID > 0]")
  # now repopulate table with selected MRA
  new_Preview_SQL = """[SQL: INSERT INTO Usr_FR_Preview (DOrder, QuestionText, AnswerList, AnswerID) 
                            SELECT DisplayOrder, QuestionText, AnswerList, -1 FROM Usr_MRA_TemplateQs WHERE TypeID = {0}]""".format(lbl_FR_Preview_ID.Content)
  _tikitResolver.Resolve(new_Preview_SQL)
  refresh_Preview_FR(s, event)
  
  ti_FR_Preview.Visibility = Visibility.Visible
  ti_FR_Overview.Visibility = Visibility.Collapsed
  
  ti_FR_Preview.IsSelected = True
  return


def Edit_FR_Template(s, event):
  # This function will load the 'Questions' tab for the selected item
  # Linked to button on XAML: btn_Edit_FRTemplate
  
  if dg_FR_Templates.SelectedIndex == -1:
    MessageBox.Show("Nothing selected to Edit!", "Error: Edit selected File Review...")
    return  
  
  # first need to load up the 'Questions' tab and then select the tab
  lbl_EditFileReview_Name.Content = dg_FR_Templates.SelectedItem['Name']
  lbl_EditFileReview_ID.Content = dg_FR_Templates.SelectedItem['Code']
  
  populate_FR_QGroups(s, event)
  refresh_FR_Questions(s, event)

  ti_FR_Questions.Visibility = Visibility.Visible
  ti_FR_Overview.Visibility = Visibility.Collapsed
  if dg_FR_Questions.Items.Count > 0:
    dg_FR_Questions.SelectedIndex = 0
  
  ti_FR_Questions.IsSelected = True
  return


def DG_FR_Template_SelectionChanged(s, event):
  # This function will populate the label controls to temp store ID and Name

  if dg_FR_Templates.SelectedIndex > -1:
    lbl_FR_Template_ID.Content = dg_FR_Templates.SelectedItem['Code']
    lbl_FR_Template_Name.Content = dg_FR_Templates.SelectedItem['Name']
  else:
    lbl_FR_Template_ID.Content = ''
    lbl_FR_Template_Name.Content = ''  
  return  
  
  
def DG_FR_Template_CellEditEnding(s, event):
  # This function will update the 'friendly name' back to the SQL table
  tmpCol = event.Column
  tmpColName = tmpCol.Header
  
  # Get Initial values updated
  updateSQL = '[SQL: UPDATE Usr_MRA_TemplateTypes SET '
  countOfUpdates = 0
  itemID = dg_FR_Templates.SelectedItem['Code']
  newName = dg_FR_Templates.SelectedItem['Name']
  newName = newName.replace("'", "''")

  if itemID != 'x':
    # Conditionally add parts depending on column updated and whether value has changed
    if tmpColName == 'Friendly Name':
      updateSQL += "TypeName = '{0}' ".format(newName) 
      countOfUpdates += 1

    # Add WHERE clause
    updateSQL += "WHERE ID = {0}]".format(itemID)
    
    # Only run if something was changed
    if countOfUpdates > 0:
      #MessageBox.Show('SQL = \n' + updateSQL)
      try:
        _tikitResolver.Resolve(updateSQL)
        refresh_FR_Templates(s, event)
      except:
        MessageBox.Show("There was an error amending the name of the File Review, using SQL:\n" + str(updateSQL), "Error: Amending Name of File Review...")
  return


class dept_Defaults(object):
  def __init__(self, myDeptN, myDepDefName, myCTGID, myRowID, myDepDefID):
     self.mraTD_DeptN = myDeptN
     self.mraTD_DeptDefaultN = myDepDefName
     self.mraTD_CTGID = myCTGID
     self.mraTD_DeptID = myRowID
     self.mraTD_DeptDefaultID = myDepDefID
     return
     
  def __getitem__(self, index):
    if index == 'DeptName':
      return self.mraTD_DeptN
    elif index == 'Default':
      return self.mraTD_DeptDefaultN
    elif index == 'CTGid':
      return self.mraTD_CTGID
    elif index == 'RowID':
      return self.mraTD_DeptID
    elif index == 'DepDefID':
      return self.mraTD_DeptDefaultID
    else:
      return ''

def refresh_FR_Department_Defaults(s, event):
  # This function will populate the DEPARTMENT Defaults datagrid (in the File Review page)
  # SQL to populate datagrid
  getTableSQL = """SELECT 'RowID' = DD.ID, 'CaseTypeGroup Name' = CTG.Name, 
                  'MRA TemplateID' = DD.TemplateID, 'MRA Template Name' = TT.TypeName, 'CaseTypeGroup ID' = CTG.ID 
                  FROM Usr_MRA_Dept_Defaults DD 
                  LEFT OUTER JOIN CaseTypeGroups CTG ON DD.CaseTypeGroupID = CTG.ID 
                  LEFT OUTER JOIN Usr_MRA_TemplateTypes TT ON DD.TemplateID = TT.TypeID 
                  WHERE DD.TypeName = 'File Review' 
                  ORDER BY CTG.Name"""
  
  tmpItem = []
  _tikitDbAccess.Open(getTableSQL)
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          tmpRowID = 0 if dr.IsDBNull(0) else dr.GetValue(0)
          tmpCTGName = '' if dr.IsDBNull(1) else dr.GetString(1)
          tmpMRATid = 0 if dr.IsDBNull(2) else dr.GetValue(2)
          tmpMRATn = '' if dr.IsDBNull(3) else dr.GetString(3)
          tmpCTGid = 0 if dr.IsDBNull(4) else dr.GetValue(4)
        
          tmpItem.append(dept_Defaults(tmpCTGName, tmpMRATn, tmpCTGid, tmpRowID, tmpMRATid))
    else:
      tmpItem.append(dept_Defaults('No items currently exist', '', 0, 0, 0))

    dr.Close()
  
  #close db connection
  _tikitDbAccess.Close()
  dg_DepartmentsFR.ItemsSource = tmpItem
  return


def FR_Department_Defaults_SelectionChanged(s, event):
  # This function will cause the Case Type list to the right to update to show selected department
  if dg_DepartmentsFR.SelectedIndex == -1:
    lbl_SelectedDeptIDFR.Content = ''
    lbl_SelectedDeptNameFR.Content = ''
    cbo_FR_Department_TemplateToUse.SelectedIndex = -1
    btn_Save_FR_TemplateToUseForDept.IsEnabled = False
    return
  
  btn_Save_FR_TemplateToUseForDept.IsEnabled = True
  
  # now refresh the case types list
  lbl_SelectedDeptIDFR.Content = dg_DepartmentsFR.SelectedItem['CTGid']
  lbl_SelectedDeptNameFR.Content = dg_DepartmentsFR.SelectedItem['DeptName']
  
  pCount = -1
  for xRow in cbo_FR_Department_TemplateToUse.Items:
    pCount += 1
    if xRow.mraT_Code == dg_DepartmentsFR.SelectedItem['DepDefID']:
      cbo_FR_Department_TemplateToUse.SelectedIndex = pCount
      break
  
  refresh_FR_CaseType_Defaults(s, event)
  return


def FR_Save_Default_For_Department(s, event):
  # This function will save the selected ID to the Department Defaults
  if len(str(lbl_SelectedDeptIDFR.Content)) == 0:
    return
  
  newTemplateID = cbo_FR_Department_TemplateToUse.SelectedItem['Code']    
  rowID = dg_DepartmentsFR.SelectedItem['RowID']
  ctgID = dg_DepartmentsFR.SelectedItem['CTGid']
  
  update_SQL = "[SQL: UPDATE Usr_MRA_Dept_Defaults SET TemplateID = {0} WHERE ID = {1}]".format(newTemplateID, rowID) 
  try: 
    _tikitResolver.Resolve(update_SQL)
  except:
    MessageBox.Show("There was an error saving the department defaults, using SQL:\n{0}".format(update_SQL), "Error: Saving Department Defaults...")
  
  refresh_FR_Department_Defaults(s, event)
  
  # This has now updated the selected item... do we also want to set all CaseTypes too?
  
  if chk_FR_ApplyToAllCaseTypes.IsChecked == True:
    updateCT_SQL = """[SQL: UPDATE Usr_MRA_CaseType_Defaults SET TemplateID = {0} WHERE TypeName = 'File Review' 
                            AND CaseTypeID IN (SELECT CT.Code FROM CaseTypes CT LEFT OUTER JOIN CaseTypeGroups CTG ON CT.CaseTypeGroupRef = CTG.ID 
                            WHERE CTG.ID = {1} AND CT.Description NOT LIKE '%Project%')]""".format(newTemplateID, ctgID)
    
    try:
      _tikitResolver.Resolve(updateCT_SQL)
    except:
      MessageBox.Show("There was an error saving the default to the Case Types, using SQL:\n{0}".format(updateCT_SQL), "Error: Saving Department Defaults (for all Case Types)...")
      return
    
    # also refresh Case Types Defaults list...
    refresh_FR_CaseType_Defaults(s, event)  
  
  return
  

class caseType_Defaults(object):
  def __init__(self, myCTname, myTemplateName, myRowID, myTemplateID, myCTid, myCTGname, myCTGID):
    self.mraT_CTName = myCTname
    self.mraT_CTTemplateToUse = myTemplateName
    self.mraT_RowID = myRowID
    self.mraT_TemplateID = myTemplateID
    self.mraT_CTid = myCTid
    self.mraT_CTGName = myCTGname
    self.mraT_CTGID = myCTGID
    return
    
  def __getitem__(self, index):
    if index == 'CTName':
      return self.mraT_CTName
    elif index == 'CTID':
      return self.mraT_CTid
    elif index == 'CTGName':
      return self.mraT_CTGName
    elif index == 'CTGid':
      return self.mraT_CTGID
    elif index == 'RowID':
      return self.mraT_RowID
    elif index == 'TemplateName':
      return self.mraT_CTTemplateToUse
    elif index == 'TemplateID':
      return self.mraT_TemplateID
    else:
      return ''


def refresh_FR_CaseType_Defaults(s, event):
  # This function will populate the 'Case Types' datagrid (for selecting which File Review template to be used)
  #! New 04/09/2025: Added SQL to add any 'missing' caseTypes that we don't have in our 'CaseTypeDefaults' table
  add_Missing_CaseTypeDefaults(forWhat='File Review')


  getTableSQL = """SELECT 'RowID' = CTD.ID, 'CaseType Name' = CT.Description, 'MRA TemplateID' = CTD.TemplateID, 
                  'MRA Template Name' = TT.TypeName, 'CaseType ID' = CT.Code, 'CaseTypeGroup Name' = CTG.Name, CT.CaseTypeGroupRef 
                  FROM Usr_MRA_CaseType_Defaults CTD 
                  LEFT OUTER JOIN CaseTypes CT ON CTD.CaseTypeID = CT.Code 
                  LEFT OUTER JOIN CaseTypeGroups CTG ON CT.CaseTypeGroupRef = CTG.ID 
                  LEFT OUTER JOIN Usr_MRA_TemplateTypes TT ON CTD.TemplateID = TT.TypeID """

  if dg_DepartmentsFR.SelectedIndex > -1:
    getTableSQL += "WHERE CTG.ID = {0} AND CTD.TypeName = 'File Review' ".format(lbl_SelectedDeptIDFR.Content)
  else:
    getTableSQL += "WHERE CTD.TypeName = 'File Review' "

  getTableSQL += "ORDER BY CT.Description"
  
  tmpItem = []
  _tikitDbAccess.Open(getTableSQL)
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          tmpRowID = 0 if dr.IsDBNull(0) else dr.GetValue(0)
          tmpCaseTypeName = '' if dr.IsDBNull(1) else dr.GetString(1)
          tmpTemplateID = 0 if dr.IsDBNull(2) else dr.GetValue(2)
          tmpTemplateName = '' if dr.IsDBNull(3) else dr.GetString(3)
          tmpCaseTypeID = 0 if dr.IsDBNull(4) else dr.GetValue(4)
          tmpCaseTypeGroupName = '' if dr.IsDBNull(5) else dr.GetString(5)
          tmpCTGID = 0 if dr.IsDBNull(6) else dr.GetValue(6)
          
          tmpItem.append(caseType_Defaults(tmpCaseTypeName, tmpTemplateName, tmpRowID, tmpTemplateID, tmpCaseTypeID, tmpCaseTypeGroupName, tmpCTGID))


    dr.Close()
  
  #close db connection
  _tikitDbAccess.Close()
  dg_FR_CaseTypes_FRTemplate.ItemsSource = tmpItem
  return


def FR_CaseType_Defaults_SelectionChanged(s, event):
  # This function populates the controls beneath the data grid to allow for updates to be made
  if dg_FR_CaseTypes_FRTemplate.SelectedIndex == -1:
    lbl_SelectedCaseTypeFR.Content = ''
    lbl_SelectedCaseTypeIDFR.Content = ''
    cbo_FR_CaseType_TemplateToUse.SelectedIndex = -1
    btn_Save_FR_TemplateToUseForCaseType.IsEnabled = False
    return

  btn_Save_FR_TemplateToUseForCaseType.IsEnabled = True
  lbl_SelectedCaseTypeFR.Content = dg_FR_CaseTypes_FRTemplate.SelectedItem['CTName']
  lbl_SelectedCaseTypeIDFR.Content = dg_FR_CaseTypes_FRTemplate.SelectedItem['CTID']
  
  pCount = -1
  for xRow in cbo_FR_CaseType_TemplateToUse.Items:
    pCount += 1
    if xRow.mraT_Code == dg_FR_CaseTypes_FRTemplate.SelectedItem['TemplateID']:
      cbo_FR_CaseType_TemplateToUse.SelectedIndex = pCount
      break
  
  return


def FR_Save_Default_For_CaseType(s, event):
  # This function saves the selected template to the Case Type defaults table
  if len(str(lbl_SelectedCaseTypeIDFR.Content)) == 0:
    return
    
  newTemplateID = cbo_FR_CaseType_TemplateToUse.SelectedItem['Code']
  rowID = dg_FR_CaseTypes_FRTemplate.SelectedItem['RowID']
  
  update_SQL = "[SQL: UPDATE Usr_MRA_CaseType_Defaults SET TemplateID = {0} WHERE ID = {1}]".format(newTemplateID, rowID)
  _tikitResolver.Resolve(update_SQL)
  
  refresh_FR_CaseType_Defaults(s, event)  
  return


# # # #    *E D I T I N G   F I L E   R E V I E W   -   Q U E S T I O N S*    TAB   # # # #

class FR_Questions(object):
  def __init__(self, myID, myDO, myQText, myQGroup, myDefCA, myQGroupID, myAllowNA, myTriggerCA, myAllowNotes):
    self.fr_ID = myID
    self.fr_DO = myDO
    self.fr_QText = myQText
    self.fr_Section = myQGroup
    self.fr_DefCA = myDefCA
    self.fr_GroupID = myQGroupID
    self.fr_AllowNA = myAllowNA
    self.fr_CAtrigger = myTriggerCA
    self.fr_AllowNotes = myAllowNotes
    return

  def __getitem__(self, index):
    if index == 'ID':
      return self.fr_ID
    elif index == 'Order':
      return self.fr_DO
    elif index == 'QText':
      return self.fr_QText
    elif index == 'Section':
      return self.fr_Section
    elif index == 'DefCA':
      return self.fr_DefCA
    elif index == 'SectionID':
      return self.fr_GroupID
    elif index == 'AllowNA':
      return self.fr_AllowNA
    elif index == 'TriggerCA':
      return self.fr_CAtrigger
    elif index == 'AllowNotes':
      return self.fr_AllowNotes
    else:
      return ''
  

def refresh_FR_Questions(s, event):
  # This function will populate the QUESTIONS datagrid On the 'Configure File Reviews > Editing Questions' tab
  # Linked XAML control: dg_FR_Questions

  getTableSQL = """SELECT '0-Order' = MRATQ.DisplayOrder, '1-QText' = MRATQ.QuestionText, '2-QGroup' = QG.Name,  
                      		'3-DefaultCA' = MRATQ.FR_Default_Corrective_Action, '4-QID' = MRATQ.QuestionID, '5-QGroupID' = QG.ID, 
                          '6-AllowNA' = ISNULL(MRATQ.FR_Allow_NA_Answer, 'Y'), '7-TriggerCA' = ISNULL(MRATQ.FR_CorrAction_Trigger_Answer, 'No'), 
                          '8-AllowNotes'  = ISNULL(MRATQ.FR_Allow_Comment, 'N')
                   FROM Usr_MRA_TemplateQs MRATQ 
	                 LEFT OUTER JOIN Usr_MRA_QGroups QG ON MRATQ.QGroupID = QG.ID
                   WHERE MRATQ.TypeID = {0} ORDER BY QG.DisplayOrder, MRATQ.DisplayOrder""".format(lbl_EditFileReview_ID.Content)

  
  tmpItem = []
  _tikitDbAccess.Open(getTableSQL)
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          tmpDO = 0 if dr.IsDBNull(0) else dr.GetValue(0)
          tmpQText = '' if dr.IsDBNull(1) else dr.GetString(1)
          tmpTabGroup = '' if dr.IsDBNull(2) else dr.GetString(2)
          tmpDefCA = '' if dr.IsDBNull(3) else dr.GetString(3)
          tmpQID = 0 if dr.IsDBNull(4) else dr.GetValue(4)
          tmpTabGroupID = 0 if dr.IsDBNull(5) else dr.GetValue(5)
          tmpAllowNA = 'Y' if dr.IsDBNull(6) else dr.GetString(6)
          tmpTriggerCA = 'No' if dr.IsDBNull(7) else dr.GetString(7)
          tmpAllowNotes = 'N' if dr.IsDBNull(8) else dr.GetString(8)
          
          #tmpItem.append(MRA_Questions(tmpQID, tmpDO, tmpQText, tmpTabGroup, tmpDefCA, tmpTabGroupID))
          tmpItem.append(FR_Questions(myID=tmpQID, myDO=tmpDO, myQText=tmpQText, myQGroup=tmpTabGroup, myDefCA=tmpDefCA, 
                                      myQGroupID=tmpTabGroupID, myAllowNA=tmpAllowNA, myTriggerCA=tmpTriggerCA, myAllowNotes=tmpAllowNotes))
    dr.Close()
  #close db connection
  _tikitDbAccess.Close()

  tmpC = ListCollectionView(tmpItem)
  tmpC.GroupDescriptions.Add(PropertyGroupDescription("fr_Section"))
  dg_FR_Questions.ItemsSource = tmpC

  if dg_FR_Questions.Items.Count > 0:
    tb_NoQuestions_FR.Visibility = Visibility.Hidden
    dg_FR_Questions.Visibility = Visibility.Visible
  else:
    tb_NoQuestions_FR.Visibility = Visibility.Visible
    dg_FR_Questions.Visibility = Visibility.Hidden
  return


def AddNew_FR_Question(s, event):
  # This function will add a new Question row to the Questions datagraid
  # Linked to XAML control.event: 
  # THIS NEEDS TO BE UPDATED : to take into account 'Groups' and therefore setting 'Display Order' accordingly

  newDO = _tikitResolver.Resolve("[SQL: SELECT ISNULL(MAX(DisplayOrder) + 1, 1) FROM Usr_MRA_TemplateQs WHERE TypeID = {0}]".format(lbl_EditFileReview_ID.Content))
  #! Added 29/07/2025: Get new QuestionID so we can pass directly into INSERT statement
  newQuestionID = _tikitResolver.Resolve("[SQL: SELECT ISNULL(MAX(QuestionID) + 1, 1) FROM Usr_MRA_TemplateQs]".format(lbl_EditFileReview_ID.Content))
  insert_SQL = """[SQL: INSERT INTO Usr_MRA_TemplateQs (TypeID, DisplayOrder, QuestionText, QuestionID, FR_Allow_NA_Answer, FR_CorrAction_Trigger_Answer, FR_Allow_Comment) 
                    VALUES({0}, {1}, '(new_question)', {2}, 'Y', 'No', 'N')]""".format(lbl_EditFileReview_ID.Content, newDO, newQuestionID)
  runSQL(codeToRun=insert_SQL, showError=True, 
         errorMsgText="There was an error adding a new question - please try again!\n\nIf issue persists, please contact IT Support", 
         errorMsgTitle="Error: Adding new File Review Question")
  # also update 'QuestionID'
  #runSQL("UPDATE Usr_MRA_TemplateQs SET QuestionID = ID WHERE QuestionID IS NULL", False, "", "")
  #! ^ nope - this is a BAD idea as we can end up using a duplicate QuestionID!

  # auto select last item...
  refresh_FR_Questions(s, event)
  dg_FR_Questions.SelectedIndex = (dg_FR_Questions.Items.Count - 1)
  txt_QuestionText_FR.Focus()
  return
  
  
def Duplicate_FR_Question(s, event):
  # This function will duplicate the selected Question and select this duped row so that it's available for editing
  if dg_FR_Questions.SelectedIndex == -1:
    MessageBox.Show("Nothing selected to duplicate!", "Error: Duplicating Selected Question...")
    return
  
  #! Added 29/07/2025: Get new QuestionID so we can pass directly into INSERT statement
  newQuestionID = _tikitResolver.Resolve("[SQL: SELECT ISNULL(MAX(QuestionID) + 1, 1) FROM Usr_MRA_TemplateQs]".format(lbl_EditFileReview_ID.Content))
  selectedID = dg_FR_Questions.SelectedItem['ID']
  insert_SQL = """[SQL: INSERT INTO Usr_MRA_TemplateQs (TypeID, DisplayOrder, QuestionText, AnswerList, TabGrouping, FR_Allow_NA_Answer, FR_CorrAction_Trigger_Answer, FR_Allow_Comment, QuestionID)  
                        SELECT TQ.TypeID, (SELECT MAX(TQ1.DisplayOrder) + 1 FROM Usr_MRA_TemplateQs TQ1 WHERE TQ1.TabGrouping = TQ.TabGrouping), 
                        TQ.QuestionText + ' (copy)', TQ.AnswerList, TQ.TabGrouping, TQ.FR_Allow_NA_Answer, TQ.FR_CorrAction_Trigger_Answer, TQ.FR_Allow_Comment, {0} 
                        FROM Usr_MRA_TemplateQs TQ WHERE TQ.ID = {1}]""".format(newQuestionID, selectedID)
  _tikitResolver.Resolve(insert_SQL)
  
  # auto select last item...
  refresh_FR_Questions(s, event)
  dg_FR_Questions.SelectedIndex = (dg_FR_Questions.Items.Count - 1)
  txt_QuestionText_FR.Focus()
  return
  

def MoveTop_FR_Question(s, event):
  # This function will move the selected Question to the top row (and all other items down one)
  tmpNewFocusRow = dgItem_MoveToTop(dgControl = dg_FR_Questions, 
                                    tableToUpdate = 'Usr_MRA_TemplateQs', 
                                    sqlOrderColName = 'DisplayOrder', 
                                    dgIDcolName = 'ID', 
                                    dgOrderColName = 'Order', 
                                    sqlOtherCheckCol = 'TypeID', 
                                    sqlOtherCheckValue = int(lbl_EditFileReview_ID.Content), 
                                    sqlGroupColName = 'QGroupID', 
                                    dgGroupColName = 'SectionID')
  if tmpNewFocusRow > -1:
    refresh_FR_Questions(s, event)
    dg_FR_Questions.Focus()
    dg_FR_Questions.SelectedIndex = tmpNewFocusRow
  return

def MoveUp_FR_Question(s, event):
  # This function will move the selected Question up one row (and all other items down one)
  tmpNewFocusRow = dgItem_MoveUp(dgControl = dg_FR_Questions, 
                                 tableToUpdate = 'Usr_MRA_TemplateQs', 
                                 sqlOrderColName = 'DisplayOrder', 
                                 dgIDcolName = 'ID', 
                                 dgOrderColName = 'Order', 
                                 sqlOtherCheckCol = 'TypeID', 
                                 sqlOtherCheckValue = int(lbl_EditFileReview_ID.Content), 
                                 sqlGroupColName = 'QGroupID', 
                                 dgGroupColName = 'SectionID')
  if tmpNewFocusRow > -1: 
    refresh_FR_Questions(s, event)
    dg_FR_Questions.Focus()
    dg_FR_Questions.SelectedIndex = tmpNewFocusRow
  return

def MoveDown_FR_Question(s, event):
  # This function will move the selected Question down one row (and all other items up one)
  tmpNewFocusRow = dgItem_MoveDown(dgControl = dg_FR_Questions, 
                                   tableToUpdate = 'Usr_MRA_TemplateQs', 
                                   sqlOrderColName = 'DisplayOrder', 
                                   dgIDcolName = 'ID', 
                                   dgOrderColName = 'Order', 
                                   sqlOtherCheckCol = 'TypeID', 
                                   sqlOtherCheckValue = int(lbl_EditFileReview_ID.Content), 
                                   sqlGroupColName = 'QGroupID', 
                                   dgGroupColName = 'SectionID')
  if tmpNewFocusRow > -1: 
    refresh_FR_Questions(s, event)
    dg_FR_Questions.Focus()
    dg_FR_Questions.SelectedIndex = tmpNewFocusRow  
  return

def MoveBottom_FR_Question(s, event):
  # This function will move the selected Question to the bottom row (and all other items up one)
  tmpNewFocusRow = dgItem_MoveToBottom(dgControl = dg_FR_Questions, 
                                       tableToUpdate = 'Usr_MRA_TemplateQs', 
                                       sqlOrderColName = 'DisplayOrder', 
                                       dgIDcolName = 'ID', 
                                       dgOrderColName = 'Order', 
                                       sqlOtherCheckCol = 'TypeID', 
                                       sqlOtherCheckValue = int(lbl_EditFileReview_ID.Content), 
                                       sqlGroupColName = 'QGroupID', 
                                       dgGroupColName = 'SectionID')
  if tmpNewFocusRow > -1:
    refresh_FR_Questions(s, event)
    dg_FR_Questions.Focus()
    dg_FR_Questions.SelectedIndex = tmpNewFocusRow  
  return

def Delete_FR_Question(s, event):
  # This function will delete the selected Question (after confirmation)
  tmpNewFocusRow = dgItem_DeleteSelected(dgControl = dg_FR_Questions, 
                                         tableToUpdate = 'Usr_MRA_TemplateQs', 
                                         sqlOrderColName = 'DisplayOrder', 
                                         dgIDcolName = 'ID', 
                                         dgOrderColName = 'Order',
                                         dgNameDescColName = 'QText', 
                                         sqlOtherCheckCol = 'TypeID', 
                                         sqlOtherCheckValue = int(lbl_EditFileReview_ID.Content), 
                                         sqlGroupColName = 'QGroupID', 
                                         dgGroupColName = 'SectionID')
  if tmpNewFocusRow > -1:
    refresh_FR_Questions(s, event)
    dg_FR_Questions.Focus()
    dg_FR_Questions.SelectedIndex = tmpNewFocusRow  
  return


def BackToOverview_FR_Question(s, event):
  # This function should clear the 'Questions' tab and take us back to the 'File Review Overview' tab
  ti_FR_Questions.Visibility = Visibility.Collapsed
  ti_FR_Overview.Visibility = Visibility.Visible
  ti_FR_Overview.IsSelected = True
  refresh_FR_Templates(s, event)
  return

def BackToOverview_FR_Preview(s, event):
  # This function should clear the 'Questions' tab and take us back to the 'File Review Overview' tab
  ti_FR_Preview.Visibility = Visibility.Collapsed
  ti_FR_Overview.Visibility = Visibility.Visible
  ti_FR_Overview.IsSelected = True
  return

def FR_Questions_SelectionChanged(s, event):
  # This function fires when the selected row is changed in the File Review 'edit questions' tab
  # Linked XAML control.event: dg_FR_Questions.SelectionChanged

  # if something is selected
  if dg_FR_Questions.SelectedIndex > -1:
    # populate the individual controls on the right with values from the datagrid
    txt_QuestionText_FR.Text = dg_FR_Questions.SelectedItem['QText']
    txt_DefaultCA.Text = dg_FR_Questions.SelectedItem['DefCA']
    lbl_QID_FR.Content = dg_FR_Questions.SelectedItem['ID']
    txt_FR_Order.Text = str(dg_FR_Questions.SelectedItem['Order'])
    chk_FR_InclNAoption.IsChecked = True if dg_FR_Questions.SelectedItem['AllowNA'] == 'Y' else False
    opt_FR_Yes.IsChecked = True if dg_FR_Questions.SelectedItem['TriggerCA'] == 'Yes' else False
    opt_FR_No.IsChecked = True if dg_FR_Questions.SelectedItem['TriggerCA'] == 'No' else False
    opt_FR_None.IsChecked = True if dg_FR_Questions.SelectedItem['TriggerCA'] == '' else False
    chk_FR_InclComments.IsChecked = True if dg_FR_Questions.SelectedItem['AllowNotes'] == 'Y' else False

    # finally select appropriate item from Section/Group list box
    dCount = -1
    for x in cbo_QuestionGroupFR.Items:
      dCount += 1
      if x.Name == dg_FR_Questions.SelectedItem['Section']:
        cbo_QuestionGroupFR.SelectedIndex = dCount
        break

  else:
    # nothing valid is selected, so set our text boxes (and combo box) to empty
    txt_QuestionText_FR.Text = ''
    txt_DefaultCA.Text = ''
    lbl_QID_FR.Content = ''
    cbo_QuestionGroupFR.SelectedIndex = -1
    chk_FR_InclNAoption.IsChecked = False
    opt_FR_Yes.IsChecked = False
    opt_FR_No.IsChecked = False
    chk_FR_InclComments.IsChecked = False
  return


def SaveChanges_FR_Question(s, event):
  # This function will save the changes made in the 'edit' area back to the database (and refreshes the data grid)
  #! Linked to XAML control.event: btn_SaveQuestion_FR.Click
  ## THIS NEEDS TO BE UPDATED TO ALSO UPDATE THE DISPLAY ORDER - THINKING HERE IS THAT USER COULD CHANGE THE 'GROUP', and therefore
  ## current DisplayOrder will not necessarily be correct... Also need to amend 'add new'
  questionID = lbl_QID_FR.Content


  if len(str(questionID)) > 0:
    # form update SQL
    tmpQ = str(txt_QuestionText_FR.Text)
    tmpQ = tmpQ.replace("'", "''")
    # check to see if a Section/Group has been selected... if not, alert user and return
    if cbo_QuestionGroupFR.SelectedIndex == -1:
      MessageBox.Show("You must select a 'Section / Group' to save the question to!", "Error: Saving Question details (File Review)...")
      return
    
    tmpQGroup = cbo_QuestionGroupFR.SelectedItem['Code']
    tmpDefaultCA = str(txt_DefaultCA.Text)
    tmpDefaultCA = tmpDefaultCA.replace("'", "''")
    tmpOrder = txt_FR_Order.Text
    tmpAllowNA = 'Y' if chk_FR_InclNAoption.IsChecked == True else 'N'
    if opt_FR_Yes.IsChecked == True:
      tmpCATrigger = 'Yes' 
    elif opt_FR_No.IsChecked == True:
      tmpCATrigger = 'No'
    else:
      tmpCATrigger = ''

    tmpAllowComments = 'Y' if chk_FR_InclComments.IsChecked == True else 'N'

    update_SQL = """[SQL: UPDATE Usr_MRA_TemplateQs SET QuestionText = '{0}', QGroupID = {1}, 
                          FR_Default_Corrective_Action = '{2}', DisplayOrder = {3}, 
                          FR_Allow_NA_Answer = '{4}', FR_CorrAction_Trigger_Answer = '{5}', 
                          FR_Allow_Comment = '{6}' WHERE QuestionID = {7}]""".format(tmpQ, tmpQGroup, tmpDefaultCA, tmpOrder, tmpAllowNA, 
                                                                                     tmpCATrigger, tmpAllowComments, questionID)
    
    #MessageBox.Show("SQL to update: {0}".format(update_SQL))

    try:
      _tikitResolver.Resolve(update_SQL)
      # refresh list
      refresh_FR_Questions(s, event)

      # now select the item again
      tCount = -1
      for myRow in dg_FR_Questions.Items:
        tCount += 1
        if int(myRow.fr_ID) == int(questionID):
          dg_FR_Questions.SelectedIndex = tCount
          break

    except:
      MessageBox.Show("There was an error saving Question details, using SQL:\n{0}".format(update_SQL), "Error: Saving Question details (File Review)...")
    
  else:
    MessageBox.Show("Nothing to save!", "Save Question details (File Review)...")
    
  return


def populate_FR_QGroups(s, event):
  # This function will populate the 'Section / Group' combo box on the 'Edit FR Questions' tab

  mySQL = "SELECT DisplayOrder, ID, Name FROM Usr_MRA_QGroups WHERE ForWhat = 'FR' ORDER BY DisplayOrder "
  tmpDDL = []
  _tikitDbAccess.Open(mySQL)
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          tmpDO = 0 if dr.IsDBNull(0) else dr.GetValue(0)
          tmpID = 0 if dr.IsDBNull(1) else dr.GetValue(1)
          tmpDesc = '' if dr.IsDBNull(2) else dr.GetString(2)

          tmpDDL.append(twoColList(myCode=tmpID, myName=tmpDesc)) 
    dr.Close()
  
  #close db connection
  _tikitDbAccess.Close()

  cbo_QuestionGroupFR.ItemsSource = tmpDDL
  return


# # # #   END OF:   *E D I T I N G   Q U E S T I O N S*    TAB   # # # #

# # # #   *P R E V I E W*    TAB    # # # #

class preview_FR(object):
  def __init__(self, myID, myOrder, myQuestion, myAnsList, myAnswer, myAList, myAnswerText):
    if len(myAList) > 0:
      tmpAList = myAList.split("|")
    else:
      tmpAList = []
  
    self.pvID = myID
    self.pvDO = myOrder
    self.pvQuestion = myQuestion
    self.pvAnswerList = myAnsList
    self.pvAnswer = '' if myAnswer == -1 else myAnswer
    self.pvSourceAnswers = tmpAList
    self.pvAnswerText = myAnswerText
    return
    
  def __getitem__(self, index):
    if index == 'Order':
      return self.pvDO
    elif index == 'Question':
      return self.pvQuestion
    elif index == 'AnswerList':
      return self.pvAnswerList
    elif index == 'Answer':
      return self.pvAnswer
    elif index == 'ID':
      return self.pvID
    elif index == 'AText':
      return self.pvAnswerText
    else:
      return ''
      
def refresh_Preview_FR(s, event):
  
  mySQL = """SELECT FRP.ID, FRP.DOrder, FRP.QuestionText, FRP.AnswerList, FRP.AnswerID, 
              'AList' = (SELECT STRING_AGG(AnswerText, '|') FROM Usr_MRA_TemplateAs WHERE GroupName = FRP.AnswerList), 
              'AnswerText' = (SELECT AnswerText FROM Usr_MRA_TemplateAs WHERE AnswerID = ID) 
            FROM Usr_FR_Preview FRP  ORDER BY FRP.DOrder"""

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
          iAnsList = '' if dr.IsDBNull(3) else dr.GetString(3)
          iAnswer = 0 if dr.IsDBNull(4) else dr.GetValue(4)
          iAList = '' if dr.IsDBNull(5) else dr.GetString(5)
          iAText = '' if dr.IsDBNull(6) else dr.GetString(6)
                    
          myItems.append(preview_FR(iID, iDO, iQText, iAnsList, iAnswer, iAList, iAText))  
      
    dr.Close()
  _tikitDbAccess.Close()
  
  dg_FRPreview.ItemsSource = myItems
  
  if dg_FRPreview.Items.Count == 0:
    dg_FRPreview.Visibility = Visibility.Hidden
    tb_NoFR_PreviewQs.Visibility = Visibility.Visible
  else:
    dg_FRPreview.Visibility = Visibility.Visible
    tb_NoFR_PreviewQs.Visibility = Visibility.Hidden
  return


def FR_Preview_CellEditEnding(s, event):
  
  # get current values
  rowID = dg_FRPreview.SelectedItem['ID']
  newTextVal = dg_FRPreview.SelectedItem['AText']          # Answer']
  fromAnsList = dg_FRPreview.SelectedItem['AnswerList']

  #MessageBox.Show("rowID: " + str(rowID) + "\nNewTextVal: " + str(newTextVal) + "\nFromAnsList: " + str(fromAnsList))

  # lookup answer index and score
  tmpAnsID = _tikitResolver.Resolve("[SQL: SELECT ID FROM Usr_MRA_TemplateAs WHERE GroupName = '{0}' AND AnswerText = '{1}']".format(fromAnsList, newTextVal))
  
  #MessageBox.Show("tmpAnsID: " + str(tmpAnsID))
  #return
  
  updateSQL = "[SQL: UPDATE Usr_FR_Preview SET AnswerID = {0} WHERE ID = {1}]".format(tmpAnsID, rowID)
  canContinue = False
  try:
    _tikitResolver.Resolve(updateSQL)
    canContinue = True
  except:
    MessageBox.Show("There was an error updating the answer (no updates made!), using SQL:\n" + updateSQL, "Error: FR Preview - Updating Answer...")
    
  if canContinue == True:
    currDGindex = dg_FRPreview.SelectedIndex
    
    if chk_FRPreview_AutoSelectNext.IsChecked == True:
      FR_Preview_AutoAdvance(currDGindex, s, event)
  return


def FR_Preview_AutoAdvance(currentDGindex, s, event):
  currPlusOne = currentDGindex + 1
  totalDGitems = dg_FRPreview.Items.Count
  #MessageBox.Show("Current Index: " + str(currentDGindex) + "\nPlusOne: " + str(currPlusOne) + "\nTotalDGitems: " + str(totalDGitems), "Auto-Advance to next question...")
  
  # firstly check to see if we're at the end of current list
  if currPlusOne == totalDGitems:
    dg_FRPreview.SelectedIndex = -1
  else:
    dg_FRPreview.SelectedIndex = currPlusOne
    dg_FRPreview.BeginEdit()
  return


def FR_Preview_SelectionChanged(s, event):
  if dg_FRPreview.SelectedIndex > -1:
    lbl_FRPreview_DGID.Content = dg_FRPreview.SelectedItem['ID']
    lbl_FRPreview_CurrVal.Content = dg_FRPreview.SelectedItem['AText']
    dg_FRPreview.BeginEdit()
  else:
    lbl_FRPreview_DGID.Content = ''
    lbl_FRPreview_CurrVal.Content = ''    
  return
 
# # # #   END OF:   *P R E V I E W*    TAB   # # # # 

# # # #   END OF:   C O N F I G U R E   F I L E   R E V I E W S   TAB   # # # #

# # # #   B E L O W   I S   T H E   C O D E   T H A T   C A N   B E   C O P Y   A N D   P A S T E D   A S - I S     # # # #
# # # #   T H E S E   A R E   T H E   M A S T E R S   A N D   S H O U L D   N O T   R E Q U I R E   U P D A T E S   # # # # 
# # # #                       NB: The code below only needs to exist ONCE per XAML or project                       # # # # 

# Common Parameters for these re-usable functions:
# dgControl         = Name of DataGrid control to be updating / getting other info from
# tableToUpdate     = Name of SQL table that needs to be updated/deleted from
# sqlOrderColName   = Name of the column as it appears in the SQL table for specifying the 'Order'/'DisplayOrder'
# dgIDcolName       = Text to use to obtain the 'ID' of item as stored in the data grid
# dgOrderColName    = Text to use to obtain the 'Order' ('DisplayOrder') of item as stored in the data grid
# dgNameDescColName = Text to use to obtain the 'Name' (or 'Description') of the item as stored in the data grid  (only on 'DELETE_SELECTED')
# mra_TypeID        = Set to -1 if not needed, but otherwise should be the 'TypeID' for the Matter Risk Assessment item
# GroupName         = As we have one table for 'File Review' and 'MRA' templates, we use this 'group' to make sure we only update the display order for items with this given group name
# ^ replaced in favour of the below:
# sqlOtherCheckCol   = Name of the other SQL column to check (if applicable, if not, provide empty string) 
# sqlOtherCheckValue = Value of the other SQL column to check (if applicable, if not, provide empty string)

# dgItem_DeleteSelected(dgControl, tableToUpdate, sqlOrderColName, dgIDcolName, dgOrderColName, dgNameDescColName, sqlOtherCheckCol, sqlOtherCheckValue)  #mra_TypeID)
# dgItem_MoveDown      (dgControl, tableToUpdate, sqlOrderColName, dgIDcolName, dgOrderColName, sqlOtherCheckCol, sqlOtherCheckValue)                     #mra_TypeID, GroupName)
# dgItem_MoveUp        (dgControl, tableToUpdate, sqlOrderColName, dgIDcolName, dgOrderColName, sqlOtherCheckCol, sqlOtherCheckValue)                     #mra_TypeID, GroupName)
# dgItem_MoveToTop     (dgControl, tableToUpdate, sqlOrderColName, dgIDcolName, dgOrderColName, sqlOtherCheckCol, sqlOtherCheckValue)                     #mra_TypeID, GroupName)
# dgItem_MoveToBottom  (dgControl, tableToUpdate, sqlOrderColName, dgIDcolName, dgOrderColName, sqlOtherCheckCol, sqlOtherCheckValue)                     #mra_TypeID, GroupName)

#! 02/09/2025: These really ought to be re-written - whilst they 'work', they're a bit clucky and in-efficient
#!             Could be cleverer with re-usable functions rather than copying logic into each as we currently have.
#!             For example: 'Delete' is the only unique one, but it does then do an update to existing DisplayOrder values
#!             (which is the same code on the other functions). So could have a 'UpdateDisplayOrder' function that takes parameters
#! Move Up - decrease 'DisplayOrder' by 1 of selected item, and increase the 'DisplayOrder' of the item that was previously in that position by 1
#!    Down - increase 'DisplayOrder' by 1 of selected item, and decrease the 'DisplayOrder' of the item that was previously in that position by 1
#!     Top - set 'DisplayOrder' of selected item to 1, and increase the 'DisplayOrder' of all other items that were previously above it by 1
#!  Bottom - set 'DisplayOrder' of selected item to max+1, and decrease the 'DisplayOrder' of all other items that were previously below it by 1


def dgItem_DeleteSelected(dgControl, tableToUpdate, sqlOrderColName, dgIDcolName, dgOrderColName, dgNameDescColName, 
                          sqlOtherCheckCol, sqlOtherCheckValue, sqlGroupColName = '', dgGroupColName = ''):
  # This function will DELETE a row from a given table (but asks for confirmation first).
  newIndexPos = -1

  if dgControl.SelectedIndex > -1:

    # Get seleted ID and details
    sel_ID = dgControl.SelectedItem[dgIDcolName]
    sel_order = '' if len(dgOrderColName) == 0 else dgControl.SelectedItem[dgOrderColName]
    sel_Name = dgControl.SelectedItem[dgNameDescColName]
    currentPos = dgControl.SelectedIndex
    if dgGroupColName != '':
      QGroupID = dgControl.SelectedItem[dgGroupColName]

    if int(sel_ID) > 0:
      msg = "Are you sure you want to delete the following item:\n{0}?".format(sel_Name)
  
      # Confirm with user before deletion
      myResult = MessageBox.Show(msg, 'Delete item...', MessageBoxButtons.YesNo)
  
      if myResult == DialogResult.Yes:
        # Form the SQL to delete row and execute the SQL 
        Delete_SQL = "DELETE FROM {0} WHERE ID = {1}".format(tableToUpdate, sel_ID)
        runSQL(Delete_SQL, True, "There was an error trying to delete item", "Error: Delete Selected...")
        
        # if supplied 'DislayOrder' column, then also update the 'DisplayOrder' for all other items
        if len(sqlOrderColName) > 0:
          # Form the SQL to update any current items with a higher DISPLAY ORDER and execute the SQL 
          UPDATE_SQL = "UPDATE {0} SET {1} = ({1} - 1) WHERE {1} > {2}".format(tableToUpdate, sqlOrderColName, sel_order) 
          
          if len(sqlOtherCheckCol) > 0:
            UPDATE_SQL += " AND {0} = {1}".format(sqlOtherCheckCol, sqlOtherCheckValue)

          if len(sqlGroupColName) > 0:
            UPDATE_SQL += " AND {0} = {1}".format(sqlGroupColName, QGroupID)

          runSQL(UPDATE_SQL, True, "There was an error trying to update the DisplayOrder for other items", "Error: Delete Selected (updating DisplayOrder)...")
      
        newIndexPos = (currentPos - 1) if (currentPos - 1) > -1 else 0
  return newIndexPos


def dgItem_MoveDown(dgControl, tableToUpdate, sqlOrderColName, dgIDcolName, dgOrderColName, sqlOtherCheckCol, sqlOtherCheckValue,
                    sqlGroupColName = '', dgGroupColName = ''):
  newIndexPos = -1

  # only continue if something selected
  if dgControl.SelectedIndex == -1:
    return newIndexPos
  
  selectedID = dgControl.SelectedItem[dgIDcolName]
  sel_dOrder = int(dgControl.SelectedItem[dgOrderColName])
  newIndexPos = dgControl.SelectedIndex + 1

  if len(sqlOtherCheckCol) > 0:
    WHEREandSQL = " AND {sqlFieldName} = {valueToCheck}".format(sqlFieldName=sqlOtherCheckCol, valueToCheck=sqlOtherCheckValue)
  else:
    WHEREandSQL = ""

  if len(sqlGroupColName) > 0 and len(dgGroupColName) > 0:
    QGroupID = dgControl.SelectedItem[dgGroupColName]
    WHEREandSQLgrp = " AND {grpColName} = {grpVal}".format(grpColName=sqlGroupColName, grpVal=QGroupID)
    # lookup max display order for current group
    grpMax_SQL = "SELECT ISNULL(MAX({displayOrder}), 1) FROM {sourceTable} WHERE {groupColName} = {grpID} AND {typeIDName} = {typeID}".format(displayOrder=sqlOrderColName, sourceTable=tableToUpdate, groupColName=sqlGroupColName, grpID=QGroupID, typeIDName=sqlOtherCheckCol, typeID=sqlOtherCheckValue)
    totalNoOfItems = runSQL(grpMax_SQL, True, "There was an error getting max display order for group", "Error: MoveItemToBottom (getting max display order for group)...", returnType='Int')
  else:
    WHEREandSQLgrp = ""
    totalNoOfItems = dgControl.Items.Count

  # if item is already at the bottom - alert user
  if sel_dOrder == int(totalNoOfItems):
    MessageBox.Show('This item cannot be moved any further down!')
    
  elif sel_dOrder < int(totalNoOfItems):
    # Firstly need to check if theres an item with the desired order id already
    countItems_sql = "[SQL: SELECT COUNT(ID) FROM {sourceTable} WHERE {orderColName} = {newValue} {andWhere1} {andGroup}]".format(sourceTable=tableToUpdate, orderColName=sqlOrderColName, newValue=(sel_dOrder + 1), andWhere1=WHEREandSQL, andGroup=WHEREandSQLgrp)
    itemToMove_SQL = "[SQL: SELECT TOP(1) ISNULL(ID, 0) FROM {sourceTable} WHERE {orderColName} = {newValue} {andWhere1} {andGroup}]".format(sourceTable=tableToUpdate, orderColName=sqlOrderColName, newValue=(sel_dOrder + 1), andWhere1=WHEREandSQL, andGroup=WHEREandSQLgrp)
      
    try:
      countItems = _tikitResolver.Resolve(countItems_sql)
    except:
      MessageBox.Show("There was an error getting count of items, using SQL:\n{0}".format(countItems_sql), "Error: MoveDown (count existing items)...")
      
    # if there's an item with new ID (display order + 1)
    if int(countItems) > 0:
      try:
        # get ID of item on row below
        itemToMove_newdID = int(_tikitResolver.Resolve(itemToMove_SQL))
      except:
        MessageBox.Show("There was an error getting ID of item to move, using SQL:\n{0}".format(itemToMove_SQL), "Error: MoveDown (get ID of existing item)...")        
        
      # form the SQL to move items UP
      sql_MoveUp = "UPDATE {sourceTable} SET {orderColName} = ({orderColName} - 1) WHERE ID = {newValue} {andWhere1} {andGroup}".format(sourceTable=tableToUpdate, orderColName=sqlOrderColName, newValue=itemToMove_newdID, andWhere1=WHEREandSQL, andGroup=WHEREandSQLgrp)
      
      # If there is an item with the previous number, we adjust that one here
      if itemToMove_newdID != 0:
        runSQL(sql_MoveUp, True, "There was an error trying to move items up", "Error: MoveDown (moving items up)...")

    # form SQL to finally move selected item DOWN
    sql_MoveDown = "UPDATE {sourceTable} SET {orderColName} = ({orderColName} + 1) WHERE ID = {newValue} {andWhere1} {andGroup}".format(sourceTable=tableToUpdate, orderColName=sqlOrderColName, newValue=selectedID, andWhere1=WHEREandSQL, andGroup=WHEREandSQLgrp)

    # now do actual moving item down
    runSQL(sql_MoveDown, True, "There was an error trying to move item down", "Error: MoveDown (moving item down)...")
  
  return newIndexPos


def dgItem_MoveUp(dgControl, tableToUpdate, sqlOrderColName, dgIDcolName, dgOrderColName, sqlOtherCheckCol, sqlOtherCheckValue,
                  sqlGroupColName = '', dgGroupColName = ''):
  # Whilst this code does largely 'work', there are some edge cases where it doesn't, and can leave 'gaps' in the display order
  # At some point, I need to re-visit and re-write this to be more robust

  newIndexPos = -1

  # only continue if something selected
  if dgControl.SelectedIndex == -1:
    return newIndexPos
  
  selectedID = int(dgControl.SelectedItem[dgIDcolName])
  sel_DisplayOrder = int(dgControl.SelectedItem[dgOrderColName])
  newIndexPos = dgControl.SelectedIndex - 1

  if len(sqlOtherCheckCol) > 0:
    WHEREandSQL = " AND {sqlFieldName} = {valueToCheck}".format(sqlFieldName=sqlOtherCheckCol, valueToCheck=sqlOtherCheckValue)
  else:
    WHEREandSQL = ""

  if len(sqlGroupColName) > 0 and len(dgGroupColName) > 0:
    QGroupID = dgControl.SelectedItem[dgGroupColName]
    WHEREandSQLgrp = " AND {grpColName} = {grpVal}".format(grpColName=sqlGroupColName, grpVal=QGroupID)
  else:
    WHEREandSQLgrp = ""

  # if current display order is already 1 (at the top) - alert user
  if sel_DisplayOrder == 1:
    MessageBox.Show('This item cannot be moved any further up!')
    
  elif sel_DisplayOrder > 1:
    # firstly need to check if theres an item with the desired order already (eg: DisplayOrder -1)
    countItems_sql = "[SQL: SELECT COUNT(ID) FROM {sourceTable} WHERE {orderColName} = {newValue} {andWhere1} {andGroup}]".format(sourceTable=tableToUpdate, orderColName=sqlOrderColName, newValue=(sel_DisplayOrder - 1), andWhere1=WHEREandSQL, andGroup=WHEREandSQLgrp)
    itemToMove_newIDsql = "[SQL: SELECT TOP (1) ISNULL(ID, 0) FROM {sourceTable} WHERE {orderColName} = {newValue} {andWhere1} {andGroup}]".format(sourceTable=tableToUpdate, orderColName=sqlOrderColName, newValue=(sel_DisplayOrder - 1), andWhere1=WHEREandSQL, andGroup=WHEREandSQLgrp)
        
    try:
      countItems = _tikitResolver.Resolve(countItems_sql)
    except:
      MessageBox.Show("There was an error getting count of items, using SQL:\n{0}".format(countItems_sql), "Error: MoveUp (count existing items)...")
      
    # if there's an item with new ID (display order + 1)
    if int(countItems) > 0:
      try:
        # get ID of item on row above
        itemToMove_newID = _tikitResolver.Resolve(itemToMove_newIDsql)
      except:
        MessageBox.Show("There was an error getting ID of item to move, using SQL:\n{0}".format(itemToMove_newIDsql), "Error: MoveUp (get ID of existing item)...")
        
      #MessageBox.Show("ItemToMove_newID: " + str(itemToMove_newID) + "\n\nFrom SQL:\n" + itemToMove_newIDsql + "\n\nSelectedID: " + str(selectedID) + "\nSel_DisplayOrder: " + str(sel_DisplayOrder) + "")

      # form the SQL to move items DOWN
      sql_MoveDown = "UPDATE {sourceTable} SET {orderColName} = ({orderColName} + 1) WHERE ID = {newValue} {andWhere1} {andGroup}".format(sourceTable=tableToUpdate, orderColName=sqlOrderColName, newValue=itemToMove_newID, andWhere1=WHEREandSQL, andGroup=WHEREandSQLgrp)

      # do the actual 'Move down' of other items
      if itemToMove_newID != 0:
        runSQL(sql_MoveDown, True, "There was an error trying to move items down", "Error: MoveUp (moving items down)...")

    # now do the actual moving item up
    sql_MoveUp = "UPDATE {sourceTable} SET {orderColName} = ({orderColName} - 1) WHERE ID = {newValue} {andWhere1} {andGroup}".format(sourceTable=tableToUpdate, orderColName=sqlOrderColName, newValue=selectedID, andWhere1=WHEREandSQL, andGroup=WHEREandSQLgrp)
      
    # now do actual moving item UP  
    runSQL(sql_MoveUp, True, "There was an error trying to move item up", "Error: MoveUp (moving item up)...")
        
  return newIndexPos


def dgItem_MoveToTop(dgControl, tableToUpdate, sqlOrderColName, dgIDcolName, dgOrderColName, sqlOtherCheckCol, sqlOtherCheckValue,
                     sqlGroupColName = '', dgGroupColName = ''):
  # This will move the selected item to the top of the list. .:. Needs to:
  # 1 - get ID of selected item so we can set that to 1 (so it's first in the list)
  # 2 - set everything with a Order value > than item from '1' above to OrderID + 1
  # 3 - refresh list and select first item
  
  newIndexPos = -1
  
  # Only continue if something is selected
  if dgControl.SelectedIndex > -1:
    
    selectedID = int(dgControl.SelectedItem[dgIDcolName])
    selectedOrder = int(dgControl.SelectedItem[dgOrderColName])
    if dgGroupColName != '':
      QGroupID = dgControl.SelectedItem[dgGroupColName]
    else:
      newIndexPos = 0
    
    if selectedOrder == 1:
      MessageBox.Show('This item is already at the top of the list!')
      
    elif selectedOrder > 1:
      # Get and run SQL to update everything with an Order < selectedOrder to current 'ItemOrder' + 1
      sql_MoveTop = "UPDATE {0} SET {1} = 1 WHERE ID = {2}".format(tableToUpdate, sqlOrderColName, selectedID)
      sql_MoveDown = "UPDATE {0} SET {1} = ({1} + 1) WHERE {1} < {2}".format(tableToUpdate, sqlOrderColName, selectedOrder)

      if len(sqlOtherCheckCol) > 0:
        sql_MoveTop += " AND {0} = {1}".format(sqlOtherCheckCol, sqlOtherCheckValue)
        sql_MoveDown += " AND {0} = {1}".format(sqlOtherCheckCol, sqlOtherCheckValue)
      # if 'group' provided, then append this check to the SQL too
      if len(sqlGroupColName) > 0:
        sql_MoveDown += " AND {0} = {1}".format(sqlGroupColName, QGroupID)

      # run SQL to move all other items down
      runSQL(sql_MoveDown, True, "There was an error moving items down", "Error: MoveItemToTop (moving other items down)...")  
      
      # Get and run SQL to update selected items 'Order' to 1
      runSQL(sql_MoveTop, True, "There was an error moving the item to the top", "Error: MoveItemToTop (move to top)...")
  
  return newIndexPos


def dgItem_MoveToBottom(dgControl, tableToUpdate, sqlOrderColName, dgIDcolName, dgOrderColName, sqlOtherCheckCol, sqlOtherCheckValue,
                        sqlGroupColName = '', dgGroupColName = ''):
  # This will move the selected item to the bottom of the list. .:. Needs to:
  # 1 - get ID of selected item so we can set that to...
  # 2 - count of total items in list
  # 3 - set everything with a Order value < than item from '1' above to OrderID - 1
  # 4 - refresh list and select last item
  newIndexPos = -1
  
  # only continue if something selected
  if dgControl.SelectedIndex == -1:
    return newIndexPos
  
  selectedID = int(dgControl.SelectedItem[dgIDcolName])
  selectedOrder = int(dgControl.SelectedItem[dgOrderColName])
  if dgGroupColName != '':
    QGroupID = dgControl.SelectedItem[dgGroupColName]

  #! below respects our 'Groups' 
  if sqlGroupColName != '':
    # lookup max display order for current group
    grpMax_SQL = "SELECT ISNULL(MAX({displayOrder}), 1) FROM {sourceTable} WHERE {groupColName} = {grpID} AND {typeIDName} = {typeID}".format(displayOrder=sqlOrderColName, sourceTable=tableToUpdate, groupColName=sqlGroupColName, grpID=QGroupID, typeIDName=sqlOtherCheckCol, typeID=sqlOtherCheckValue)
    totalNoOfItems = runSQL(grpMax_SQL, True, "There was an error getting max display order for group", "Error: MoveItemToBottom (getting max display order for group)...", returnType='Int')
  else:
    totalNoOfItems = dgControl.Items.Count
    newIndexPos = totalNoOfItems - 1
        
  if selectedOrder == totalNoOfItems:
    MessageBox.Show('This item is already at the bottom of the list!')
    
  elif selectedOrder < totalNoOfItems:
    # Get and run SQL to update everything with an Order > selectedOrder to current 'ItemOrder' - 1
    sql_MoveBottom = "UPDATE {sourceTable} SET {sqlField} = {newValue} WHERE ID = {idOfItem}".format(sourceTable=tableToUpdate, sqlField=sqlOrderColName, newValue=totalNoOfItems, idOfItem=selectedID)
    sql_MoveUp = "UPDATE {sourceTable} SET {sqlField} = ({sqlField} - 1) WHERE {sqlField} > {currentOrder}".format(sourceTable=tableToUpdate, sqlField=sqlOrderColName, currentOrder=selectedOrder)

    if len(sqlOtherCheckCol) > 0:
      sql_MoveBottom += " AND {sqlField} = {newValue}".format(sqlField=sqlOtherCheckCol, newValue=sqlOtherCheckValue)
      sql_MoveUp += " AND {sqlField} = {newValue}".format(sqlField=sqlOtherCheckCol, newValue=sqlOtherCheckValue)
    if len(sqlGroupColName) > 0:
      sql_MoveUp += " AND {sqlField} = {newValue}".format(sqlField=sqlGroupColName, newValue=QGroupID)

    # run SQL to move all other items up
    runSQL(sql_MoveUp, True, "There was an error moving items up", "Error: MoveItemToBottom (moving other items up)...")

    # Get and run SQL to update selected items 'Order' to last item
    runSQL(sql_MoveBottom, True, "There was an error moving the item to the bottom", "Error: MoveItemToBottom (move to bottom)...")

  return newIndexPos

# # # #   E N D   O F   S E C T I O N :  O N C E   P E R   P R O J E C T   O R   X A M L   C O D E     # # # # 

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
    countOfEntities = _tikitResolver.Resolve("[SQL: SELECT COUNT(Code) FROM Entities WHERE Code = '" + myFullLenString + "']")
    # if above count is zero, we return the short ref so other functions return 'error' otherwise we provide the full length code
    if int(countOfEntities) == 0:
      myFinalString = shortRef
    else:
      myFinalString = myFullLenString
 
    #MessageBox.Show("GetFullLenEntityRef - Input: " + str(shortRef) + "\nOutput: " + myFinalString + "\nCount of Entities: " + str(countOfEntities))
  return myFinalString


def runSQL(codeToRun = '', showError = False, errorMsgText = '', errorMsgTitle = '', 
           useAltResolver = False, returnType = 'String', textBoxForOutput = None):
  # Traditionally, we used to use _tikitResolver.Resolve() as-is, but have since found that it's better to wrap this within Python's 'try: except:' construct.
  # In order to minimise code, I made this reusable function to do so plus allow for a custom message to be displayed upon error.  See below for explanation of inputs/arguments:
  # codeToRun             = Full SQL of code to run. No need to wrap in '[SQL: code_Here]' as we can do that here
  # showError             = True | False (default). Indicates whether or not to display message box upon error
  # errorMsgText          = (empty string default). Text to display in the body of the message box upon error (note: actual SQL will automatically be included, so no need to re-supply that)
  # errorMsgTitle         = (empty string default). Text to display in the title bar of the message box upon error
  # useAltResolver        = True | False (default). Indicates whether to use the alternative resolver (eg: for stored procedures, CTEs, etc) or not
  # returnType            = 'String' (default) | 'Int' | 'DataReader'. Indicates what type of value to return from the SQL execution
  #! ^ above params added, as there are two different 'resolver' methods in use, one for normal SQL statements, and another for stored procedures, CTEs, etc.
  #!   Therefore, re-written this function to also be able to use the other with the same 'runSQL()' function
  # textBoxForOutput      = None (defatul) | TextBox. If supplied, will output the result of the SQL execution into this TextBox
  # codeToRun_returnValue = (empty string default). This expects a standard 'SELECT' statement, and should be what you'd like returned from this procedure.
  #                         Eg: your main 'codeToRun' may be an 'INSERT INTO...' and you would like the code/ID/row_number of the added item returned - with the
  #                             'codeToRun_returnValue' parameter, you can specify the SQL to retrieve such info... Generally speaking, this is usually immediately
  #                             after an 'INSERT INTO' in my code anyway, so just need to pass as a param here instead.
  #                         ^ ignoring this... seemed like a good idea, but couldn't make use off in this particular project (WIP WriteOff)
  #                           What we do here is: 1) count existing rows match entity/matter; 2) if none exist, add one; 3) get id of row for entity/matter
  #                           (therefore there's not really any 'wastage' there, and don't gain anything by having param in here [because in above example
  #                            the code to 'insert' is behind an 'if' statement, so if there were already rows that exist, we'd never get new id])
  #                           If anything, it may just make more sense to have smaller functions where we can whack these multi-line SQL's into (eg purpose specific)

  # Note: calling procedure can use like we do with '_tikitResolver()', that is: 
  # - tmpValue = runSQL("SELECT YEAR()")   # to capture value into a variable, or:
  # - runSQL("INSERT INTO x () VALUES()")  # to just run the SQL without saving to variable
  
  #! Note: if adding 'returnSQL' param, we'll want to add a check at beginning for any 'INSERT | UPDATE | DELETE' keywords to help identify when to return 'TRUE | FALSE' (to state event happened)
  #! or in case of 'SELECT'

  if textBoxForOutput is not None:
    textBoxForOutput.Text += "{0} > runSQL(showError: {2}; errorMsgText: '{3}'; errorMsgTitle: '{4}'; useAltResolver: {5}; returnType: '{6}')\ncodeToRun: {1}\n".format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 
                                                                                                                                                                        codeToRun,
                                                                                                                                                                        showError,
                                                                                                                                                                        errorMsgText,
                                                                                                                                                                        errorMsgTitle,
                                                                                                                                                                        useAltResolver,
                                                                                                                                                                        returnType)


  # if no code actually supplied, exit early
  if len(codeToRun) < 10:
    if textBoxForOutput is not None:
      textBoxForOutput.Text += "{0}> Output: Error - length of code doesn't appear long enough\n".format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    else:
      if showError == True:
        MessageBox.Show("The supplied 'codeToRun' doesn't appear long enough, please check and update this code if necessary.\nPassed SQL: " + str(codeToRun), "ERROR: runSQL...")
    return "Error"
  
  # search for 'INSERT INTO' | 'UPDATE' | 'DELETE' in passed 'code to run'...
  pattern = r'\b(?:INSERT\s+INTO|DELETE|UPDATE)\b'
  isActionSQL = bool(re.search(pattern, codeToRun, re.IGNORECASE))
  if isActionSQL == True and textBoxForOutput is not None:
    textBoxForOutput.Text += "SQL contains 'INSERT | UPDATE | DELETE' (will return 'TRUE|FALSE' indicating success instead)\n"


  if useAltResolver == False:
    # Use the standard 'resolver' for normal/simple 'SELECT' statements, or 'INSERT', 'UPDATE', 'DELETE' etc.
    # Add '[SQL: ]' wrapper if not already included
    if codeToRun[:5] == "[SQL:":
      fCodeToRun = codeToRun
    else:
      fCodeToRun = "[SQL: {0}]".format(codeToRun)
  
    # try to execute the SQL
    try:
      tmpValue = _tikitResolver.Resolve(fCodeToRun)
      didError = False

      # if a TextBox was supplied, then output the result into it
      if textBoxForOutput is not None:
        if isActionSQL == True:
          textBoxForOutput.Text += '{0} > Output: True (INSERT|UPDATE|DELETE executed successfully)\n'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        else:
          if returnType == 'Int':
            textBoxForOutput.Text += '{0} > Output: {1}\n'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"), tmpValue)
          else:
            textBoxForOutput.Text += "{0} > Output: '{1}'\n".format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"), tmpValue)

      # return an integer to calling proc if requested
      if returnType == 'Int':
        return int(tmpValue) if tmpValue is not None else 0
      else:
        return tmpValue

    except Exception as e:
      didError = True
      # there was an error... check to see if opted to show message or not...
      if showError == True:
        MessageBox.Show("{0}\nSQL used:\n{1}\nException:{2}".format(errorMsgText, codeToRun, str(e)), errorMsgTitle)
      
            # if a TextBox was supplied, then output the result into it
      if textBoxForOutput is not None:
        if isActionSQL == True:
          textBoxForOutput.Text += '{0} > Output: False (INSERT|UPDATE|DELETE did not execute)...\nException: {1}\n'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"), str(e))
        else:
          textBoxForOutput.Text += '{0} > Output: Error - Unexpected Error Occurred...\nException: {1}\n'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"), str(e))

      if returnType == 'Int':
        return 0
      else:
        return "Error"
    
  else:
    # use alternative resolver for executing stored procedures, or other complex SQL (eg: CTEs)
    didError = False

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

      # output value to text box
      if textBoxForOutput is not None:
        if isActionSQL == True:
          textBoxForOutput.Text += '{0} > Output: True (INSERT|UPDATE|DELETE executed successfully)\n'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        else:
          if returnType == 'Int':
            textBoxForOutput.Text += '{0} > Output: {1}\n'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"), tmpValue)
          else:
            textBoxForOutput.Text += "{0} > Output: '{1}'\n".format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"), tmpValue)

      return tmpValue
  
    except Exception as e:
      # if there was an error with the CTE supplied, we'll get to here, so update outputs accordingly
      didError = True

      if showError == True:
        MessageBox.Show("{0}\nSQL used:\n{1}\nException:{2}".format(errorMsgText, codeToRun, str(e)), errorMsgTitle)

      # if a TextBox was supplied, then output the result into it
      if textBoxForOutput is not None:
        if isActionSQL == True:
          textBoxForOutput.Text += '{0} > Output: False (INSERT|UPDATE|DELETE did not execute)...\nException: {1}\n'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"), str(e))
        else:
          textBoxForOutput.Text += '{0} > Output: Error - Unexpected Error Occurred...\nException: {1}\n'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"), str(e))        

      if returnType == 'Int':
        return 0
      else:
        return "Error"
    
###################################################################################################################################################

]]>
    </Init>
    <Loaded>
      <![CDATA[
#Define controls that will be used in all of the code
TC_Main = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'TC_Main')
ti_FR_Overview = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_FR_Overview')
ti_FR_Questions = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_FR_Questions')
ti_FR_Preview = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_FR_Preview')


## M A N A G E   L O C K S   -  TAB ##
cboDept = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cboDept')
cboFeeEarner = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cboFeeEarner')
txtSearch = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'txtSearch')
txtEntityRef = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'txtEntityRef')
txtMatterNo = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'txtMatterNo')
cbo_GroupBy = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cbo_GroupBy')
cbo_GroupBy.SelectionChanged += cbo_GroupBy_SelectionChanged  
btn_RefreshLockedMatters = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_RefreshLockedMatters')
btn_RefreshLockedMatters.Click += btn_RefreshLockedMatters_Click
btn_ClearLockedMattersFilters = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_ClearLockedMattersFilters')
btn_ClearLockedMattersFilters.Click += clear_LockedMatters_Filters

# Unlock Matter button controls
btn_UnlockMatters2 = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_UnlockMatters2')
btn_UnlockMatters2.Click += btn_UnlockMatters2_Click
popTemplates = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'popTemplates')
popTemplates.Closed += popTemplates_Closed
icTemplates = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'icTemplates')
btnCancel = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btnCancel')
btnCancel.Click += CancelPopup_Click

dg_LockedMatters = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_LockedMatters')
tb_NoLockedMatters = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_NoLockedMatters')


## C O N F I G U R E   F I L E   R E V I E W S   - TAB ##
btn_AddNew_FRTemplate = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_AddNew_FRTemplate')
btn_AddNew_FRTemplate.Click += AddNew_FR_Template
btn_CopySelected_FRTemplate = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_CopySelected_FRTemplate')
btn_CopySelected_FRTemplate.Click += Duplicate_FR_Template
btn_DeleteSelected_FRTemplate = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_DeleteSelected_FRTemplate')
btn_DeleteSelected_FRTemplate.Click += Delete_FR_Template
btn_Preview_FRTemplate = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Preview_FRTemplate')
btn_Preview_FRTemplate.Click += Preview_FR_Template
btn_Edit_FRTemplate = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Edit_FRTemplate')
btn_Edit_FRTemplate.Click += Edit_FR_Template

lbl_FR_Template_ID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_FR_Template_ID')
lbl_FR_Template_Name = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_FR_Template_Name')

dg_FR_Templates = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_FR_Templates')
dg_FR_Templates.SelectionChanged += DG_FR_Template_SelectionChanged
dg_FR_Templates.CellEditEnding += DG_FR_Template_CellEditEnding

dg_DepartmentsFR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_DepartmentsFR')
dg_DepartmentsFR.SelectionChanged += FR_Department_Defaults_SelectionChanged
dg_FR_CaseTypes_FRTemplate = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_FR_CaseTypes_FRTemplate')
dg_FR_CaseTypes_FRTemplate.SelectionChanged += FR_CaseType_Defaults_SelectionChanged
lbl_SelectedDeptNameFR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_SelectedDeptNameFR')
lbl_SelectedDeptIDFR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_SelectedDeptIDFR')
cbo_FR_Department_TemplateToUse = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cbo_FR_Department_TemplateToUse')
btn_Save_FR_TemplateToUseForDept = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Save_FR_TemplateToUseForDept')
btn_Save_FR_TemplateToUseForDept.Click += FR_Save_Default_For_Department
lbl_SelectedCaseTypeFR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_SelectedCaseTypeFR')
lbl_SelectedCaseTypeIDFR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_SelectedCaseTypeIDFR')
cbo_FR_CaseType_TemplateToUse = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cbo_FR_CaseType_TemplateToUse')
btn_Save_FR_TemplateToUseForCaseType = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Save_FR_TemplateToUseForCaseType')
btn_Save_FR_TemplateToUseForCaseType.Click += FR_Save_Default_For_CaseType
chk_FR_ApplyToAllCaseTypes = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'chk_FR_ApplyToAllCaseTypes')

lbl_EditFileReview_ID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_EditFileReview_ID')
lbl_EditFileReview_Name = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_EditFileReview_Name')
btn_AddNew_Q_FR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_AddNew_Q_FR')
btn_AddNew_Q_FR.Click += AddNew_FR_Question
btn_CopySelected_Q_FR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_CopySelected_Q_FR')
btn_CopySelected_Q_FR.Click += Duplicate_FR_Question
btn_Q_MoveTop_FR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Q_MoveTop_FR')
btn_Q_MoveTop_FR.Click += MoveTop_FR_Question
btn_Q_MoveUp_FR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Q_MoveUp_FR')
btn_Q_MoveUp_FR.Click += MoveUp_FR_Question
btn_Q_MoveDown_FR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Q_MoveDown_FR')
btn_Q_MoveDown_FR.Click += MoveDown_FR_Question
btn_Q_MoveBottom_FR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Q_MoveBottom_FR')
btn_Q_MoveBottom_FR.Click += MoveBottom_FR_Question
btn_DeleteSelected_Q_FR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_DeleteSelected_Q_FR')
btn_DeleteSelected_Q_FR.Click += Delete_FR_Question
btn_BackToOverview_FR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_BackToOverview_FR')
btn_BackToOverview_FR.Click += BackToOverview_FR_Question

dg_FR_Questions = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_FR_Questions')
dg_FR_Questions.SelectionChanged += FR_Questions_SelectionChanged
btn_SaveQuestion_FR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_SaveQuestion_FR')
btn_SaveQuestion_FR.Click += SaveChanges_FR_Question
lbl_QID_FR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_QID_FR')
txt_QuestionText_FR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'txt_QuestionText_FR')
tb_NoQuestions_FR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_NoQuestions_FR')
# new
txt_DefaultCA = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'txt_DefaultCA')
cbo_QuestionGroupFR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cbo_QuestionGroupFR')
txt_FR_Order = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'txt_FR_Order')
opt_FR_Yes = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'opt_FR_Yes')
opt_FR_No = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'opt_FR_No')
opt_FR_NA = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'opt_FR_NA')
opt_FR_None = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'opt_FR_None')
chk_FR_InclNAoption = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'chk_FR_InclNAoption')
chk_FR_InclComments = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'chk_FR_InclComments')


## P R E V I E W   F I L E   R E V I E W   - TAB ##
btn_FRPreview_BackToOverview = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_FRPreview_BackToOverview')
btn_FRPreview_BackToOverview.Click += BackToOverview_FR_Preview
lbl_FR_Preview_ID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_FR_Preview_ID')
lbl_FR_Preview_Name = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_FR_Preview_Name')
tb_NoFR_PreviewQs = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_NoFR_PreviewQs')
dg_FRPreview = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_FRPreview')
dg_FRPreview.CellEditEnding += FR_Preview_CellEditEnding
dg_FRPreview.SelectionChanged += FR_Preview_SelectionChanged
lbl_FRPreview_DGID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_FRPreview_DGID')
lbl_FRPreview_CurrVal = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_FRPreview_CurrVal')
chk_FRPreview_AutoSelectNext = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'chk_FRPreview_AutoSelectNext')


## M A N A G E   A N S W E R S   - TAB ##
btn_AddNewList = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_AddNewList')
btn_AddNewList.Click += addNewList
btn_CopySelectedList = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_CopySelectedList')
btn_CopySelectedList.Click += duplicateSelectedList
btn_DeleteSelectedList = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_DeleteSelectedList')
#btn_DeleteSelectedList.Click += deleteSelectedList
lbl_List_SelID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_List_SelID')
lbl_List_SelDesc = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_List_SelDesc')
lbl_NoGlobalGroups = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_NoGlobalGroups')
dg_Lists = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_Lists')
dg_Lists.SelectionChanged += ListGroup_SelectionChanged
# now added a prompt into function below to confirm updating of Global List name
dg_Lists.CellEditEnding += ListGroup_CellEditEnding
btn_RefreshAnswerList = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_RefreshAnswerList')
btn_RefreshAnswerList.Click += refresh_AnswerListGroups

lbl_SelectedList = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_SelectedList')
lbl_SelectedListQID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_SelectedListQID')
btn_AddNewListItem = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_AddNewListItem')
btn_AddNewListItem.Click += addNewListItem
btn_CopySelectedListItem = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_CopySelectedListItem')
btn_CopySelectedListItem.Click += duplicateSelectedListItem
btn_A_MoveTop = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_A_MoveTop')
btn_A_MoveTop.Click += Answers_MoveItemToTop
btn_A_MoveUp = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_A_MoveUp')
btn_A_MoveUp.Click += Answers_MoveItemUp
btn_A_MoveDown = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_A_MoveDown')
btn_A_MoveDown.Click += Answers_MoveItemDown
btn_A_MoveBottom = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_A_MoveBottom')
btn_A_MoveBottom.Click += Answers_MoveItemToBottom
btn_DeleteSelectedListItem = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_DeleteSelectedListItem')
btn_DeleteSelectedListItem.Click += deleteSelectedListItem
dg_ListItems = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_ListItems')
dg_ListItems.SelectionChanged += ListItems_SelectionChanged
dg_ListItems.CellEditEnding += ListItems_CellEditEnding

tb_NoAnswerItems = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_NoAnswerItems')

lbl_ListItemText = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_ListItemText')
lbl_ListItemScore = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_ListItemScore')
lbl_ListItemEC = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_ListItemEC')

## New GROUPS Section (on Manage Answers tab)
lbl_GroupItemText = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_GroupItemText')
lbl_GroupItemGroup = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_GroupItemGroup')
tb_NoGroupItems = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_NoGroupItems') 
dg_GroupItems = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_GroupItems')
dg_GroupItems.SelectionChanged += Group_SelectionChanged
dg_GroupItems.CellEditEnding += Group_CellEditEnding
btn_AddNewGroupItem = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_AddNewGroupItem')
btn_AddNewGroupItem.Click += addNewGroup
btn_AddNewGroupItemFR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_AddNewGroupItemFR')
btn_AddNewGroupItemFR.Click += addNewGroupFR
btn_DeleteSelectedGroupItem = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_DeleteSelectedGroupItem')
btn_DeleteSelectedGroupItem.Click += deleteSelectedGroup
btn_CopySelectedGroupItem = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_CopySelectedGroupItem')
btn_CopySelectedGroupItem.Click += duplicateSelectedGroup
#btn_Group_MoveTop = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Group_MoveTop')
#btn_Group_MoveTop.Click += Group_MoveItemToTop
btn_Group_MoveUp = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Group_MoveUp')
btn_Group_MoveUp.Click += Group_MoveItemUp
btn_Group_MoveDown = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Group_MoveDown')
btn_Group_MoveDown.Click += Group_MoveItemDown
#btn_Group_MoveBottom = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Group_MoveBottom')
#btn_Group_MoveBottom.Click += Group_MoveItemToBottom

# Define Actions and on load events
myOnLoadEvent(_tikitSender, 'onLoad')

]]>
    </Loaded>
  </RiskPractice_v2>

</tfb>