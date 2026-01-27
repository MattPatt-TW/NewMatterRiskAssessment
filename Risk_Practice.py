<?xml version="1.0" encoding="UTF-8"?>
<tfb>
  <RiskPractice>
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
from System.Windows import Controls, Forms, LogicalTreeHelper
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
  refresh_ListOfLockedMatters(s, event)       # Locked Matters refresh
  refresh_MRA_Templates(s, event)             # Matter Risk Assessment (main overview / templates) 
  refresh_MRA_Department_Defaults(s, event)   # MRA Default Template for Department
  refresh_MRA_CaseType_Defaults(s, event)     # MRA Default Template for Case Type
  refresh_FR_Templates(s, event)              # File Review (main overview / templates)
  refresh_FR_Department_Defaults(s, event)    # FR Default Template for Department
  refresh_FR_CaseType_Defaults(s, event)      # FR Default Template for Case Type
  refresh_AnswerListGroups(s, event)          # Answers list (shared for both MRA and FR)
  refresh_GroupItems(s, event)                # New 'Group' items (for MRA 'Section/Group')

  set_Visibility_ofAnswerItemsDG()  
  
  # hide 'Edit Questions' and 'Preview' tabs
  ti_MRA_Questions.Visibility = Visibility.Collapsed
  ti_MRA_Preview.Visibility = Visibility.Collapsed
  ti_FR_Questions.Visibility = Visibility.Collapsed
  ti_FR_Preview.Visibility = Visibility.Collapsed
  
  # hide other controls not needed until something is selected...
  grd_ScoreThresholds.Visibility = Visibility.Collapsed
  stk_ST_SelectedMRA.Visibility = Visibility.Collapsed
  btn_SaveScoreThresholds.Visibility = Visibility.Collapsed
  tb_ST_NoMRA_Selected.Visibility = Visibility.Visible
  btn_Save_MRA_TemplateToUseForDept.IsEnabled = True
  cbo_Reason.ItemsSource = [
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
  #MessageBox.Show("Hello world! - OnLoad Finished")
  return

# # # #   L O C K E D   M A T T E R S   # # # #

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
  def __init__(self, myOurRef, myCLName, myMatDesc, myFE, myDept, myMO, myDSMO, myEntRef, myMatNo, myMRAID, myMRAName, myMRAExpiry, myTTExpDays):
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
    else:
      return ''
      
def refresh_ListOfLockedMatters(s, event):

  # first need to get the lock ID for 'LockedByRiskDepartment' lock
  lockID_SQL = """SELECT CASE WHEN EXISTS(SELECT Code FROM Locks WHERE Description = 'LockedByRiskDept') THEN 
                  (SELECT Code FROM Locks WHERE Description = 'LockedByRiskDept') ELSE 0 END """
  lockID = runSQL(lockID_SQL, False, '', '')
  
  if int(lockID) == 0:
    MessageBox.Show("There doesn't appear to be a Lock setup in the name of 'LockedByRiskDept', so cannot list matters locked with this lock!", "Error: Refresh List of Locked Matters...")
    return
  
  # This function will populate the list of locked matters (first main tab)
  mySQL = """SELECT E.ShortCode + '/' + CONVERT(nvarchar, M.Number), E.LegalName, M.Description, 
            'Fee Earner' = '(' + U.Code + ') ' + U.FullName, 'Dept' = CTG.Name, M.Created, 
            'Days since matter open' = DATEDIFF(DAY, M.Created, GETDATE()), M.EntityRef, M.Number, 
            'OV ID' = MRAO.ID, 'OV Name' = MRAO.LocalName, 'ExpiryDate' = MRAO.ExpiryDate, 'TT ExpiryDays' = TT.ValidityPeriodDays 
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
          myItems.append(MatterLocks(iOurRef, iClName, iMatDesc, iFE, iDept, iMO, iDSMO, iEntRef, iMatNo, iMRAid, iMRAName, iMRAExp, iTTExpDays))  

      
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
    else:
      dpItem.append('-NA-')
      
    dr.Close()
  else:
    dpItem.append('-NA-')
  
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
  
  refresh_ListOfLockedMatters(s, event)
  return
  

def unlockMatter(s, event):
  # This function will unlock the selected matter(s)
  # AND (as of 29/02/2024) will also extend the 'Expiry Date' for all MRA items by their default time
  # Note, still needs updating to allow selection of many items (button implies we can select many rows and currently do allow multi-select on DG (was 'Extended' now set to 'Single')
  if dg_LockedMatters.SelectedIndex == -1:
    MessageBox.Show("No matter has been selected!\nPlease select a matter before clicking 'Unlock'", "Error: Unlock Selected Matter...")
    return
  if cbo_Reason.SelectedItem is None:
    MessageBox.Show("No reason for the matter locking has been selected")
    return
  # get entityRef and MatterNo and form SQL to run the stored procedure
  tmpOurRef = dg_LockedMatters.SelectedItem['OurRef']
  tmpEntity = dg_LockedMatters.SelectedItem['EntRef']
  tmpMatter = dg_LockedMatters.SelectedItem['MatNo']
  tmpOVID = dg_LockedMatters.SelectedItem['MRAID']
  tmpEmailTo = runSQL("SELECT EMailExternal FROM Users WHERE Code = (SELECT FeeEarnerRef FROM Matters WHERE EntityRef = '{0}' AND Number = {1})".format(tmpEntity, tmpMatter), False, '', '')
  tmpEmailCC = ''
  tmpToUserName = runSQL("SELECT ISNULL(Forename, FullName) FROM Users WHERE Code = (SELECT FeeEarnerRef FROM Matters WHERE EntityRef = '{0}' AND Number = {1})".format(tmpEntity, tmpMatter), False, '', '')
  tmpMatDesc = dg_LockedMatters.SelectedItem['MatterDesc']
  tmpClName = dg_LockedMatters.SelectedItem['ClientName']
  tmpAddtl1 = dg_LockedMatters.SelectedItem['MRA Name']
  tmpAddtl2 = runSQL("SELECT CASE RiskRating WHEN 1 THEN 'Low' WHEN 2 THEN 'Medium' WHEN 3 THEN 'High' ELSE '-unknown-' END FROM Usr_MRA_Overview WHERE ID = {0}".format(tmpOVID), False, '', '')
  tmpExpDays = dg_LockedMatters.SelectedItem['TT ExpDays']

  unlockCode = "[SQL: EXEC TW_LockHandler '{0}', {1}, 'LockedByRiskDept', 'UnLock']".format(tmpEntity, tmpMatter)
    
  update_log_SQL = "INSERT INTO Usr_Unlock_Log (MatterNo, EntityRef, Reason, Date_Unlocked) VALUES ({0}, '{1}', '{2}', GETDATE())".format(tmpMatter, tmpEntity, cbo_Reason.SelectedItem)
  runSQL(update_log_SQL, True, "There was an error updating the unlock log", "Error: Unlock Selected Matter - Updating Log...")

  cbo_Reason.SelectedIndex = -1

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
  refresh_ListOfLockedMatters(s, event)
  return
  

# # # #   END OF:   L O C K E D   M A T T E R S   # # # # 


# # # #   M A T T E R   R I S K   A S S E S S M E N T   T E M P L A T E S   # # # #

class MRA_Templates(object):
  def __init__(self, myCode, myFName, myCountUsed, myQCount, myExpiryDays, myInternalNote, 
               myUsersNote, myHidden, myEffFrom, myEffTo, myEditingTypeID, myRowID, myParentID, myVersionNo, myIsPublished):
    self.mraT_Code = myCode
    self.mraT_Desc = myFName
    self.CountUsed = myCountUsed
    self.QCount = myQCount
    self.mraT_VPeriod = myExpiryDays
    self.MRAInternalNote = myInternalNote
    self.MRAUsersNote = myUsersNote
    self.MRAHidden = myHidden
    self.MRAFrom = myEffFrom
    self.MRATo = myEffTo
    self.MRAEditingTypeID = myEditingTypeID
    self.MRA_InEditMode = 'N' if myEditingTypeID == -1 else 'Y'
    self.MRA_ID = myRowID
    self.NRA_ParentID = myParentID
    self.NRA_VersionNo = myVersionNo
    self.NRA_IsPublished = myIsPublished
    return

  def __getitem__(self, index):
    if index == 'Code':
      # Note: this is the TypeID
      return self.mraT_Code
    elif index == 'ID':
      # Note: this is the actual unique 'ID' (row ID)
      return self.MRA_ID
    elif index == 'Name':
      return self.mraT_Desc
    elif index == 'CountUsed':
      return self.CountUsed
    elif index == 'QCount':
      return self.QCount
    elif index == 'ExpiryDays':
      return self.mraT_VPeriod
    elif index == 'InternalNote':
      return self.MRAInternalNote
    elif index == 'UserNote':
      return self.MRAUsersNote
    elif index == 'Hidden':
      return self.MRAHidden
    elif index == 'FromDate':
      return self.MRAFrom
    elif index == 'ToDate':
      return self.MRATo
    elif index == 'EditingTypeID':
      return self.MRAEditingTypeID
    elif index == 'InEditMode':
      return self.MRA_InEditMode
    elif index == 'ParentID':
      return self.NRA_ParentID
    elif index == 'VersionNo':
      return self.NRA_VersionNo
    elif index == 'IsPublished':
      return self.NRA_IsPublished
    else:
      return ''
  

def refresh_MRA_Templates(s, event):
  # This funtion populates the main Matter Risk Assessment data grid (and also populates the combo drop-downs in the 'Department' and 'Case Type' defaults area)
  
  # SQL to populate datagrid
  getTableSQL = """SELECT MRA_TT.TypeID, MRA_TT.TypeName, 
                          'Count Used' = (SELECT COUNT(ID) FROM Usr_MRA_Overview WHERE TypeID = MRA_TT.TypeID), 
                          'QCount' = (SELECT COUNT(ID) FROM Usr_MRA_TemplateQs TQs WHERE TQs.TypeID = MRA_TT.TypeID), 
                          'Expiry' =  MRA_TT.ValidityPeriodDays,
                          MRA_TT.InternalNote, MRA_TT.UsersNote, MRA_TT.Hidden, 
                          MRA_TT.EffectiveFrom, MRA_TT.EffectiveTo, MRA_TT.EditingTypeID, MRA_TT.ID, 
                          MRA_TT.ParentTypeID, MRA_TT.VersionNo, MRA_TT.IsPublished
                  FROM Usr_MRA_TemplateTypes MRA_TT WHERE MRA_TT.Is_MRA = 'Y' """
  
  if chk_ShowHiddenNMRAtemplates.IsChecked == False:
    getTableSQL += "AND MRA_TT.Hidden = 'N' "
  getTableSQL += "ORDER BY MRA_TT.TypeID"

  tmpItem = []
  
  _tikitDbAccess.Open(getTableSQL)
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          tmpTypeID = 0 if dr.IsDBNull(0) else dr.GetValue(0)               
          tmpName = '' if dr.IsDBNull(1) else dr.GetString(1)
          tmpCU = 0 if dr.IsDBNull(2) else dr.GetValue(2)
          tmpQCount = 0 if dr.IsDBNull(3) else dr.GetValue(3)
          tmpVDays = 0 if dr.IsDBNull(4) else dr.GetValue(4)
          tmpInNote = '' if dr.IsDBNull(5) else dr.GetString(5)
          tmpUserNote = '' if dr.IsDBNull(6) else dr.GetString(6)
          tmpHidden = '' if dr.IsDBNull(7) else dr.GetString(7)
          tmpFromD = 0 if dr.IsDBNull(8) else dr.GetValue(8) 
          tmpToD = 0 if dr.IsDBNull(9) else dr.GetValue(9)
          tmpETI = -1 if dr.IsDBNull(10) else dr.GetValue(10)
          tmpID = 0 if dr.IsDBNull(11) else dr.GetValue(11)
          tmpParentID = 0 if dr.IsDBNull(12) else dr.GetValue(12)
          tmpVersionNo = 0 if dr.IsDBNull(13) else dr.GetValue(13)
          tmpIsPublished = '' if dr.IsDBNull(14) else dr.GetString(14)

          tmpItem.append(MRA_Templates(myCode=tmpTypeID, myFName=tmpName, myCountUsed=tmpCU, myQCount=tmpQCount, 
                                       myExpiryDays=tmpVDays, myInternalNote=tmpInNote, myUsersNote=tmpUserNote,
                                        myHidden=tmpHidden, myEffFrom=tmpFromD, myEffTo=tmpToD, myEditingTypeID=tmpETI, myRowID=tmpID,
                                        myParentID=tmpParentID, myVersionNo=tmpVersionNo, myIsPublished=tmpIsPublished))
    dr.Close()
  #close db connection
  _tikitDbAccess.Close()
  dg_MRA_Templates.ItemsSource = tmpItem
  return


def DG_MRA_Template_SelectionChanged(s, event):
  # This function will populate the label controls to temp store ID and Name
  #! updated 26/08/2025 to include the other new fields we added

  if dg_MRA_Templates.SelectedIndex > -1:
    selItem = dg_MRA_Templates.SelectedItem
    tb_Sel_MRA_Name.Text = str(selItem['Name'])
    lbl_Sel_MRA_ID.Content = str(selItem['Code'])
    lbl_Sel_MRA_EditingTypeID.Content = '-N/A-' if str(selItem['EditingTypeID']) == '' else str(selItem['EditingTypeID'])

    #lbl_Sel_MRA_Name.Content = str(selItem['Name'])   # < this is the 'copy of' name (so may need to look this up?? Or just Ignore??)
    tb_ExpiresInXdays.Text = str(selItem['ExpiryDays']) if selItem['ExpiryDays'] is not None else '0'

    tb_InternalNote.Text = str(selItem['InternalNote'])
    tb_UsersNote.Text = str(selItem['UserNote'])
    chk_MRA_Hidden.IsChecked = True if selItem['Hidden'] == 'Y' else False
    dtp_MRA_EffectiveFrom.SelectedDate = selItem['FromDate'] if isinstance(selItem['FromDate'], DateTime) else None
    dtp_MRA_EffectiveTo.SelectedDate = selItem['ToDate'] if isinstance(selItem['ToDate'], DateTime) else None

    stk_ST_SelectedMRA.Visibility = Visibility.Visible
    grd_ScoreThresholds.Visibility = Visibility.Visible
    btn_SaveScoreThresholds.Visibility = Visibility.Visible
    tb_ST_NoMRA_Selected.Visibility = Visibility.Collapsed
    refresh_MRA_LMH_ScoreThresholds_New(s, event)
    
    btn_CopySelected_MRATemplate.IsEnabled = True
    btn_Preview_MRATemplate.IsEnabled = True
    btn_Edit_MRATemplate.IsEnabled = True
    btn_DeleteSelected_MRATemplate.IsEnabled = True if int(dg_MRA_Templates.SelectedItem['CountUsed']) == 0 else False
    btn_Save_Main_MRA_Header_Details.IsEnabled = True
  else:

    tb_Sel_MRA_Name.Text = ''
    lbl_Sel_MRA_ID.Content = ''

    tb_ExpiresInXdays.Text = '0'
    tb_InternalNote.Text = ''
    tb_UsersNote.Text = ''
    chk_MRA_Hidden.IsChecked = False
    dtp_MRA_EffectiveFrom.SelectedDate = None
    dtp_MRA_EffectiveTo.SelectedDate = None

    stk_ST_SelectedMRA.Visibility = Visibility.Collapsed
    grd_ScoreThresholds.Visibility = Visibility.Collapsed
    btn_SaveScoreThresholds.Visibility = Visibility.Collapsed
    tb_ST_NoMRA_Selected.Visibility = Visibility.Visible
    
    btn_CopySelected_MRATemplate.IsEnabled = False
    btn_Preview_MRATemplate.IsEnabled = False
    btn_Edit_MRATemplate.IsEnabled = False
    btn_DeleteSelected_MRATemplate.IsEnabled = False
    btn_Save_Main_MRA_Header_Details.IsEnabled = False 
  return


def btn_SaveMainMRA_Details_Click(s, event):
  # This is the main 'Save' button on the 'List of NMRA Templates' tab, and now replaces the 'CellEditEnding' event

  if dg_MRA_Templates.SelectedIndex == -1:
    MessageBox.Show("No Matter Risk Assessment Template has been selected!\nPlease select a template before clicking 'Save Changes'", "Error: Save Changes to Selected Matter Risk Assessment Template...")
    return

  itemID = lbl_Sel_MRA_ID.Content
  newName = str(tb_Sel_MRA_Name.Text)
  newName = newName.replace("'", "''")

  try:
    newExpDays = int(tb_ExpiresInXdays.Text)
  except:
    MessageBox.Show("The 'Expires in ?? days' value must be a whole number (integer)", "Error: Save Changes to Selected NMRA Template...")
    return
  
  newInNote = str(tb_InternalNote.Text)
  newInNote = newInNote.replace("'", "''")
  newUserNote = str(tb_UsersNote.Text)
  newUserNote = newUserNote.replace("'", "''")

  # form the SQL to update
  updateSQL = """UPDATE Usr_MRA_TemplateTypes SET TypeName = '{0}', ValidityPeriodDays = {1}, InternalNote = '{2}', UsersNote = '{3}',
                 Hidden = '{4}' WHERE TypeID = {5} """.format(newName, newExpDays, newInNote, newUserNote, 
                                                              'Y' if chk_MRA_Hidden.IsChecked == True else 'N', itemID)

  # do update
  try:
    _tikitResolver.Resolve("[SQL: " + updateSQL + "]")
    refresh_MRA_Templates(s, event)

    # re-select the same item in the data grid by iterating over list and matching ID
    tCount = -1
    for x in dg_MRA_Templates.Items:
      tCount += 1
      if str(x.mraT_Code) == str(itemID):
        dg_MRA_Templates.SelectedIndex = tCount
        break

    MessageBox.Show("Successfully updated details of selected Matter Risk Assessment Template", "Save Changes to Selected NMRA Template - Success...")
  except:
    MessageBox.Show("There was an error amending the details of the Matter Risk Assessment Template, using SQL:\n" + str(updateSQL), "Error: Amending Details of Matter Risk Assessment Template...")

  return


def DG_MRA_Template_CellEditEnding(s, event):
  # This function will update the 'friendly name' back to the SQL table
  tmpCol = event.Column
  tmpColName = tmpCol.Header
  
  # Get Initial values updated
  updateSQL = "[SQL: UPDATE Usr_MRA_TemplateTypes SET "
  countOfUpdates = 0
  itemID = dg_MRA_Templates.SelectedItem['Code']
  newName = dg_MRA_Templates.SelectedItem['Name']
  newName = newName.replace("'", "''")
  newExpDays = dg_MRA_Templates.SelectedItem['ExpiryDays']

  if itemID != 'x':
    # Conditionally add parts depending on column updated and whether value has changed
    if tmpColName == 'Friendly Name':
      updateSQL += "TypeName = '{0}' ".format(newName)
      countOfUpdates += 1
    
    if tmpColName == 'Expires in ?? days':
      updateSQL += "ValidityPeriodDays = {0} ".format(newExpDays)
      countOfUpdates += 1
    
    # Add WHERE clause
    updateSQL += "WHERE TypeID = {0}]".format(itemID)
    
    # Only run if something was changed
    if countOfUpdates > 0:
      #MessageBox.Show('SQL = \n' + updateSQL)
      try:
        _tikitResolver.Resolve(updateSQL)
        refresh_MRA_Templates(s, event)
      except:
        MessageBox.Show("There was an error amending the name of the Matter Risk Assessment, using SQL:\n" + str(updateSQL), "Error: Amending Name of Matter Risk Assessment...")
  return


def AddNew_MRA_Template(s, event):
  # This function will add a new row to the 'Matter Risk Assessments' data drid
  
  # firstly, create a uniquie name, as we'll need this to get the ID later ('MRA - [DATE]' in format YYYY-MM-DD HH:MM:SS)
  tmpName = _tikitResolver.Resolve("[SQL: SELECT 'Matter Risk Assessment - ' + CONVERT(nvarchar, GETDATE(), 120)]")
  #! Added 29/07/2025: Get next new TypeID so we can add it directly in the INSERT statement
  nextTypeID = _tikitResolver.Resolve("[SQL: SELECT ISNULL(MAX(TypeID), 0) + 1 FROM Usr_MRA_TemplateTypes]")
  insertSQL = "[SQL: INSERT INTO Usr_MRA_TemplateTypes (TypeName, Is_MRA, TypeID) VALUES ('{0}', 'Y', {1})]".format(tmpName, nextTypeID)
  
  try:
    _tikitResolver.Resolve(insertSQL)
  except:
    MessageBox.Show("There was an error trying to create a new Matter Risk Assessment, using SQL:\n" + str(insertSQL), "Error: Adding new Matter Risk Assessment...")
    return
  
  ## now get ID of added item...
  #newMRA_SQL = "[SQL: SELECT ID FROM Usr_MRA_TemplateTypes WHERE TypeName = '{0}']".format(tmpName)
  #try:
  #  newMRA_ID = _tikitResolver.Resolve(newMRA_SQL)
  #except:
  #  newMRA_ID = 0
  #  MessageBox.Show("There was an error trying to get ID of newly added Matter Risk Assessment, using SQL:\n" + str(newMRA_SQL), "Error: Adding new Matter Risk Assessment...")
  #  return
  # 
  # if int(newMRA_ID) > 0:
  #   # Also update the 'TypeID'
  #   runSQL("UPDATE Usr_MRA_TemplateTypes SET TypeID = ID WHERE ID = {0}".format(newMRA_ID), False, '', '')
  #! nope - this is a BAD idea as we could end up duplicating TypeID, so instead, we get next number up front and add directly in the INSERT statement above

  # Now create a new 'Score Threshold' for newly added item
  addNew_ScoreThreshold(nextTypeID)
  # refresh 'Score Matrix' area... (might occur automatically upon selecting new item from data grid next)
  
  # refresh data grid and select last item
  refresh_MRA_Templates(s, event)
  dg_MRA_Templates.Focus()
  #dg_MRA_Templates.SelectedIndex = (dg_MRA_Templates.Items.Count - 1)    # < don't want to assume last item is newly added item, as we know name, iterate over and select accordingly
  
  tCount = -1
  for x in dg_MRA_Templates.Items:
    tCount += 1
    if str(x.mraT_Desc) == str(tmpName):
      dg_MRA_Templates.SelectedIndex = tCount
      break
  return


class QATidyUp(object):
  def __init__(self, myNewID, mySourceID, myQuestionID):
    self.NewID = myNewID
    self.SourceID = mySourceID
    self.QuestionID = myQuestionID
    return

  def __getitem__(self, index):
    if index == 'NewID':
      return self.NewID
    elif index == 'SourceID':
      return self.SourceID
    elif index == 'QuestionID':
      return self.QuestionID
    else:
      return ''

def btn_Duplicate_MRA_Template(s, event):
  # This function will duplicate the selected Matter Risk Assessment (including the questions) AND ANSWERS - NEED TO REVIEW
  #! 19/08/2025 - Need to make this into its own dedicated function with parameters, as will be use on 'Edit' button too (to copy selected MRA template)

  if dg_MRA_Templates.SelectedIndex == -1:
    MessageBox.Show("Nothing selected to copy!", "Error: Duplicate Selected Matter Risk Assessment...")
    return
  
  idItemToCopy = dg_MRA_Templates.SelectedItem['Code']
  nameToCopy = dg_MRA_Templates.SelectedItem['Name']
  
  # Firstly, copy main template (Usr_MRA_TemplateTypes) and get new ID
  tempName = " (copy of {0})".format(idItemToCopy)
  duplicatedID = duplicate_MRA_Template(sourceTypeID=idItemToCopy, newTypeName=tempName)
  if duplicatedID == -1:
    MessageBox.Show("Error duplicating '{0}' (TypeID: {1})".format(nameToCopy, idItemToCopy), "Error: Duplicating Matter Risk Assessment...")
    return


  refresh_MRA_Templates(s, event)

  dg_MRA_Templates.Focus()
  #dg_MRA_Templates.SelectedIndex = (dg_MRA_Templates.Items.Count - 1)      # < don't want to assume last item is newly added item, as we know name, iterate over and select accordingly
  
  tCount = -1
  for x in dg_MRA_Templates.Items:
    tCount += 1
    if str(x['ID']) == str(duplicatedID):   # NB: if this doesn't work, then try direct name instead: MRA_ID (or x.MRA_ID)
      dg_MRA_Templates.SelectedIndex = tCount
      dg_MRA_Templates.ScrollIntoView(dg_MRA_Templates.Items[tCount])
      break

  MessageBox.Show("Successfully copied '{0}' as '{1}'".format(nameToCopy, tempName), "Success: Duplicating Matter Risk Assessment...")
  return


def Delete_MRA_Template(s, event):
  #! TODO: 30/07/2025: This needs to be re-written to do 'soft delete' (mark as 'Hidden' and hide from view)
  # This function will delete the selected Matter Risk Assessment template (and any questions associated to it)
  
  if dg_MRA_Templates.SelectedIndex == -1:
    MessageBox.Show("Nothing selected to delete!", "Error: Delete Selected Matter Risk Assessment...")
    return

  # First get the ID, as we'll also want to delete questions using this ID
  tTypeID = dg_MRA_Templates.SelectedItem['Code'] 

  #! NEED A 'Please Confirm' type prompt here... AND: need to check if this MRA is in use on any matters (if so, can't delete unless we re-assign)
  #! NB: 'DeleteSelected' function DOES ask for confirmation

  # Call generic function to do main delete
  tmpNewFocusRow = dgItem_DeleteSelected(dgControl=dg_MRA_Templates, 
                                         tableToUpdate='Usr_MRA_TemplateTypes', 
                                         sqlOrderColName='', 
                                         dgIDcolName='ID',
                                         dgOrderColName='', 
                                         dgNameDescColName='Name', 
                                         sqlOtherCheckCol='', 
                                         sqlOtherCheckValue='')
  if tmpNewFocusRow > -1:
    refresh_MRA_Templates(s, event)
    dg_MRA_Templates.Focus()
    dg_MRA_Templates.SelectedIndex = tmpNewFocusRow
    
    # now to delete all ANSWERS associated to questions with this this ID, then delete the QUESTIONS
    deleteA_SQL = "DELETE FROM Usr_MRA_TemplateAs WHERE QuestionID IN (SELECT QuestionID FROM Usr_MRA_TemplateQs WHERE TypeID = {0})".format(tTypeID)
    runSQL(deleteA_SQL, True, "There was an error deleting the Answers associated to the Questions used for the selected Matter Risk Assessment", "Error: Deleting Matter Risk Assessment Template...")    
    deleteQ_SQL = "DELETE FROM Usr_MRA_TemplateQs WHERE TypeID = {0}".format(tTypeID)
    runSQL(deleteQ_SQL, True, "There was an error deleting the Questions associated to the selected Matter Risk Assessment", "Error: Deleting Matter Risk Assessment Template...")
    # remove 'EditingTypeID' using this now deleted typeID
    upSQL = "UPDATE Usr_MRA_TemplateTypes SET EditingTypeID = NULL WHERE EditingTypeID = {typeIDbeingDeleted}".format(typeIDbeingDeleted=tTypeID)
    runSQL(upSQL, True, "There was an error removing the EditingTypeID from any existing Template.\n\nUsing SQL:\n{sqlRan}".format(sqlRan=upSQL), "Error: Deleting Matter Risk Assessment Template...")
    # finally the 'ScoreMatrix' items too
    delSM_sql = "DELETE FROM Usr_MRA_ScoreMatrix WHERE TypeID = {0}".format(tTypeID)
    runSQL(delSM_sql, True, "There was an error deleting the associated Score Thresholds for this Template.\n\nUsing SQL:\n{0}".format(delSM_sql), "Error: Deleting Matter Risk Assessment Template...")
  return
  
  
def Preview_MRA_Template(s, event):
  # This function will load the 'Preview' tab (made to look like 'matter-level' XAML) for the selected item
  if dg_MRA_Templates.SelectedIndex == -1:
    MessageBox.Show("Nothing selected to Preview!", "Error: Preview selected Matter Risk Assessment...")
    return  
  
  MRA_Code = dg_MRA_Templates.SelectedItem['Code']
  # first need to load up questions onto tab and then display/select tab
  lbl_MRA_Preview_ID.Content = str(MRA_Code)
  lbl_MRA_Preview_Name.Content = dg_MRA_Templates.SelectedItem['Name']
  #MessageBox.Show("MRA Code: " + str(MRA_Code) + "\nName: " + dg_MRA_Templates.SelectedItem['Name'], "DEBUG - Preview MRA Template")
  
  # clear existing list
  runSQL(codeToRun="DELETE FROM Usr_MRA_Preview WHERE ID > 0", showError=True, errorMsgText="There was an error clearing the 'Preview' table", errorMsgTitle="DEBUG - Preview MRA Template")
  
  # get count of current items (if no questions - we can't preview)
  countOfQs = _tikitResolver.Resolve("[SQL: SELECT COUNT(ID) FROM Usr_MRA_TemplateQs WHERE TypeID = {0}]".format(MRA_Code))
  
  if int(countOfQs) == 0:
    MessageBox.Show("No questions exist for the selected MRA, nothing to preview!", "Error: Previewing Matter Risk Assessment...")
    return
  
  # now repopulate table with selected MRA
  new_Preview_SQL = """[SQL: INSERT INTO Usr_MRA_Preview (DOrder, QuestionText, AnswerList, AnswerID, Score, QGroupID, EmailComment, QuestionID) 
                        SELECT DisplayOrder, QuestionText, AnswerList, -1, 0, QGroupID, '', QuestionID 
                        FROM Usr_MRA_TemplateQs WHERE TypeID = {0}]""".format(MRA_Code)

  try:
    _tikitResolver.Resolve(new_Preview_SQL)
  except:
    MessageBox.Show("There was an error attempting to preview the selected MRA, using SQL:\n" + str(new_Preview_SQL), "Error: Previewing Matter Risk Assessment...")
    return
  
  try:
    populate_preview_MRA_QGroups(s, event)
  except:
    MessageBox.Show("There was an error calling 'populate_preview_MRA_QGroups'", "DEBUG - Preview MRA Template")
  
  #refresh_Preview_MRA(s, event)
  # Don't need to call this because changing the selected Group item (which we do next) causes it to refresh then
    
  dg_GroupItems_Preview.SelectedIndex = 1
  dg_MRAPreview.SelectedIndex = 0
  
  ti_MRA_Preview.Visibility = Visibility.Visible
  ti_MRA_Overview.Visibility = Visibility.Collapsed
  MRA_Preview_UpdateTotalScore(s, event)
  
  ti_MRA_Preview.IsSelected = True
  
  return
  

def duplicate_MRA_Template(sourceTypeID=0, newTypeName=''):
  #! Added 09/01/2026 - New central function to replace multiple instances elsewhere (in 'btn_Edit_MRATemplate_Click' and 'btn_Duplicate_MRA_Template')
  # Note: we're removing the '[Editing]' text from the name, and using new fields added 'VersionNo' and 'IsPublished' to manage versions instead

  if sourceTypeID == 0:
    MessageBox.Show("No source TypeID provided to duplicate!", "Error: Duplicating Matter Risk Assessment...")
    return -1

  #! Added 29/07/2025: Get next new TypeID so we can add it directly in the INSERT statement
  nextTypeID = _tikitResolver.Resolve("[SQL: SELECT ISNULL(MAX(TypeID), 0) + 1 FROM Usr_MRA_TemplateTypes]")
  #! Get the next QuestionID and next AnswerID, so that we can pass directly into the INSERT statement (to copy over questions/answers)
  #! NB: don't +1 here as 'RowNum' starts at 1 and will mean that we skip a number each time
  nextQID = _tikitResolver.Resolve("[SQL: SELECT ISNULL(MAX(QuestionID), 0) FROM Usr_MRA_TemplateQs]")
  nextAnsID = _tikitResolver.Resolve("[SQL: SELECT ISNULL(MAX(AnswerID), 0) FROM Usr_MRA_TemplateAs]")

  # Insert main 'header' row items for the new template
  insertSQL = """INSERT INTO Usr_MRA_TemplateTypes (TypeName, Is_MRA, ValidityPeriodDays, TypeID, Hidden, EffectiveFrom, EffectiveTo, InternalNote, UsersNote, IsPublished) 
                 SELECT CONCAT(TypeName, '{newNameAppendage}'), 'Y', ValidityPeriodDays, {newTypeID}, 'Y', NULL, '2100-01-01 00:00:00.000', '', '', 'N'
                 FROM Usr_MRA_TemplateTypes WHERE TypeID = {idTOcopy}""".format(newTypeID=nextTypeID, idTOcopy=sourceTypeID, newNameAppendage=newTypeName)
  try:
    _tikitResolver.Resolve(insertSQL)
    addedTemplateMRA = True
  except Exception as e:
    MessageBox.Show("There was an error duplicating the selected Matter Risk Assessment:\n{error} using SQL:\n{errSQL}".format(error=str(e), errSQL=insertSQL), "Error: Duplicating Matter Risk Assessment...")
    addedTemplateMRA = False
    return -1
  
  # Next, we need to copy over the Questions (Usr_MRA_TemplateQs) and Answers (Usr_MRA_TemplateAs) for this NEW template
  # and we'll need new ID's and set the TypeID to the new one we just created above
  copyQ_SQL = """WITH myCopyList AS (
                      SELECT 'qText' = QuestionText, 'aList' = AnswerList, 'qGroupID' = QGroupID, 'tTypeID' = TypeID, 
                             'qID' = QuestionID,	
                             'RowNumInGrp' = ROW_NUMBER() OVER (PARTITION BY QGroupID ORDER BY DisplayOrder, QuestionText),
                             'RowNum' = ROW_NUMBER() OVER (ORDER BY QGroupID, DisplayOrder, QuestionText)
                      FROM Usr_MRA_TemplateQs WHERE TypeID = {TypeID_to_copy}
	                  )
                   INSERT INTO Usr_MRA_TemplateQs (QuestionText, AnswerList, QGroupID, TypeID, DisplayOrder, QuestionID, SourceID)
                   SELECT qText, aList, qGroupID, {newTypeID}, RowNumInGrp, {nextQID} + RowNum, qID
                   FROM myCopyList ORDER BY qGroupID, RowNumInGrp;""".format(TypeID_to_copy=sourceTypeID, newTypeID=nextTypeID, nextQID=nextQID)  
    
  didAddQs = runSQL(codeToRun=copyQ_SQL.strip(), showError=False, useAltResolver=True)

  if didAddQs == 'Error':
    addedQs = False
    MessageBox.Show("An error occurred copying the Questions, using SQL:\n" + str(copyQ_SQL), "Error: Duplicating Matter Risk Assessment...")
    return -1
  else:
    addedQs = True

  # COPYING ANSWERS OVER
  #! 30/07/2025 - New version that doesn't need loop process
  copyAnsSQL = """WITH myAnswerList AS (
                    SELECT 'GrpName' = TA.GroupName, 'AnsText' = TA.AnswerText, 'Score' = TA.Score, 
                          'EmailComment' = TA.EmailComment, 
                          'RowNum' = ROW_NUMBER() OVER (ORDER BY TQ.DisplayOrder, TA.DisplayOrder, TA.AnswerText), 
                          'AOrder' = TA.DisplayOrder, 'QOrder' = TQ.DisplayOrder, 'OrigQID' = TQ.QuestionID 
                    FROM Usr_MRA_TemplateAs TA 
                      JOIN Usr_MRA_TemplateQs TQ ON TA.QuestionID = TQ.QuestionID 
                    WHERE TQ.TypeID = {typeIDtoCopy} 
                  ), myQuestionList AS (
                    SELECT 'QOrder' = TQ.DisplayOrder, 'OrigQID' = TQ.QuestionID, 
                            'RowNumQ' = ROW_NUMBER() OVER (ORDER BY QG.DisplayOrder, TQ.DisplayOrder)
                        FROM Usr_MRA_TemplateQs TQ 
                            JOIN Usr_MRA_QGroups QG ON TQ.QGroupID = QG.ID
                    WHERE TQ.TypeID = {typeIDtoCopy} 
                  ) 
                  INSERT INTO Usr_MRA_TemplateAs (GroupName, QuestionID, AnswerText, Score, EmailComment, DisplayOrder, AnswerID) 
                  SELECT myA.GrpName, {newQNum} + myQ.RowNumQ, myA.AnsText, myA.Score, myA.EmailComment, myA.AOrder, {nextAnswerID} + myA.RowNum 
                  FROM myAnswerList myA 
                    JOIN myQuestionList myQ ON myA.OrigQID = myQ.OrigQID 
                  ORDER BY myA.RowNum;
                  """.format(typeIDtoCopy=sourceTypeID, nextAnswerID=nextAnsID, newQNum=nextQID)

  # ^ this should now be 'fixed', as we've separated Q and A, so Qnum only advances when original Qnum changes and our Answers advance for every row
  didAddAs = runSQL(codeToRun=copyAnsSQL.strip(), showError=False, useAltResolver=True)

  if didAddAs == 'Error':
    addedAs = False
    MessageBox.Show("An error occurred copying the Answers, using SQL:\n" + str(copyAnsSQL), "Error: Duplicating Matter Risk Assessment...")
    return -1
  else:
    addedAs = True

  # finally remove SourceID now we've obtained answers (don't do per Question - just do all at end)
  #updateSQL = "[SQL: UPDATE Usr_MRA_TemplateQs SET SourceID = null, QuestionID = ID WHERE ID = {0}]".format(tmpNewID)
  #MessageBox.Show("updateSQL: " + str(updateSQL), "Duplicate Matter Risk Assessment - Answers for Questions...")
  #_tikitResolver.Resolve("[SQL: UPDATE Usr_MRA_TemplateQs SET SourceID = NULL WHERE SourceID IS NOT NULL]")

  if addedTemplateMRA == True and addedQs == True and addedAs == True:
    # finally, also duplicate the Score Matrix for the MRA
    duplicate_ScoreThresholds(TypeID_toCopy=sourceTypeID, newTypeID=nextTypeID)
    #MessageBox.Show("Successfully duplicated Matter Risk Assessment Template '{0}' as new Template '{1}'".format(sourceTypeID, nextTypeID), "Success: Duplicating Matter Risk Assessment...")
    return nextTypeID
  else:
    return -1


def load_MRA_Template_ForEditing(TypeIDtoUse=None, originalItemID=0):
  # This function will load the 'Editing MRA Template' tab for the specified TypeID

  # #   L O A D   T H E   'E D I T I N G   M R A   T E M P L A T E'   T A B # #
  # first need to load up the 'Questions' tab and then select the tab

  if TypeIDtoUse is not None and TypeIDtoUse != '':

    tb_ThisMRAid.Text = str(TypeIDtoUse)
    # lookup name of 'new' (editing) template
    tLName = runSQL("SELECT TOP 1 TypeName FROM Usr_MRA_TemplateTypes WHERE TypeID = {0}".format(TypeIDtoUse))
    tb_ThisMRAname.Text = str(tLName)

    if originalItemID == 0:
      tb_CopyOfMRAid.Text = str(TypeIDtoUse)
      tb_CopyOfMRAname.Text = str(tLName)
    else:
      # this means user has selected to 'show hidden', and then chosen to click 'Edit' on the '[Editing]' item
      # so we need to now lookup those details - but can we be sure we're picking up correct original template?
      tID = runSQL("SELECT TOP 1 TypeID FROM Usr_MRA_TemplateTypes WHERE EditingTypeID = {0}".format(TypeIDtoUse))
      tName = runSQL("SELECT TOP 1 TypeName FROM Usr_MRA_TemplateTypes WHERE EditingTypeID = {0}".format(TypeIDtoUse))
      tb_CopyOfMRAid.Text = tID
      tb_CopyOfMRAname.Text = tName
      
  else:
    MessageBox.Show("There was an error trying to edit this item - please copy to IT Support for resolution.\n\nOriginal TypeID: {0}\nEditing TypeID: {1}".format(origItem['Code'], TypeIDtoUse), "Error: Editing Matter Risk Assessment...")
    return
  
  refresh_MRA_Questions(for_MRA_ID=TypeIDtoUse)
  populate_MRA_QGroups()

  ti_MRA_Questions.Visibility = Visibility.Visible
  ti_MRA_Overview.Visibility = Visibility.Collapsed
  #MRA_Questions_SelectionChanged(s, event)
  if dg_MRA_Questions.Items.Count > 0:
    dg_MRA_Questions.SelectedIndex = 0
  
  ti_MRA_Questions.IsSelected = True

  return


def btn_Edit_MRATemplate_Click(s, event):
  # This function will load the 'Questions' tab for the selected item
  #! 20/08/2025 - This is being re-written so that we never 'edit' a template that could currently be in use, but instead, we 'copy' the template and edit this version
  #!              First, checks 'EditingTypeID' column for current row and if not null, will load up and use the 'editing' version of the template, else, if null, 
  #!              this means there aren't any pending 'edits', so we will copy the selected template and set the 'EditingTypeID' to this new 'TypeID'
  #!              This will allow us to 'edit' the template without affecting the current version in use, and then we can 'save' the new version when done
  #!              (NB: there's a new 'Publish' button that when clicked, will overwrite the 'CaseTypeDefaults' table with the new template TypeID, and the old template will be 'hidden' from view) 
  #MessageBox.Show("EditSelected_Click", "DEBUG - TESTING")
  #! 09/01/2026 - NEED TP UPDATE THIS THIS FUNCTION TO CALL NEW 'duplicate_MRA_Template' FUNCTION to reduce code duplication
  # Clicking 'Edit' button on a selected NMRA template:
  # - if already an 'editing' version (EditingTypeID not null), load this one
  # - if not, duplicate selected template and set 'EditingTypeID' to new TypeID, then load this one
  
  # if nothing selected, alert user and bomb-out now...
  if dg_MRA_Templates.SelectedIndex == -1:
    MessageBox.Show("Nothing selected to Edit!", "Error: Edit selected Matter Risk Assessment...")
    return

  origItem = dg_MRA_Templates.SelectedItem 
  origItemID = origItem['Code']
  #MessageBox.Show("Original Item ID: {0}".format(origItem['Code']), "DEBUG - TESTING")
  #! NB: don't want to be using/relying on 'Name' containing '[Editing]' as this could be changed by user - instead, use 'IsPublished' and 'VersionNo' fields going forward
  tmpEditingTypeID = origItem['EditingTypeID'] if origItem['EditingTypeID'] is not None else -1
  #nameContainsEditing = True if '[Editing]' in str(origItem['Name']) else False
  isPublished = origItem['IsPublished'] if origItem['IsPublished'] is not None else 'Y'

  # firstly check to see if there's an 'EditingTypeID' set for this item - if so, this is the one we'll edit
  if tmpEditingTypeID != -1 and isPublished == 'N':
    # load 'edit' page with this TypeID
    load_MRA_Template_ForEditing(TypeIDtoUse=tmpEditingTypeID, originalItemID=origItemID)
    return

  # else if no editingTypeID and currently marked as 'published', then we need to duplicate the selected template for editing
  if isPublished == 'Y':
    #! Template is currently published, so we need to duplicate it for editing
    duplicatedID = duplicate_MRA_Template(sourceTypeID=origItemID, newTypeName='')

    if duplicatedID == -1:
      MessageBox.Show("Error duplicating '{0}' (TypeID: {1}) for editing".format(origItem['Name'], origItemID), "Error: Editing Matter Risk Assessment...")
      return
    
    #! now update the original item to set the 'EditingTypeID' to this new TypeID
    updateEditingTypeID_SQL = "[SQL: UPDATE Usr_MRA_TemplateTypes SET EditingTypeID = {0} WHERE TypeID = {1}]".format(duplicatedID, origItemID)
    didUpdateEditingTypeID = runSQL(codeToRun=updateEditingTypeID_SQL, showError=False, useAltResolver=False)
    if didUpdateEditingTypeID == 'Error':
     MessageBox.Show("There was an error updating the 'EditingTypeID' for the original Matter Risk Assessment, using SQL:\n{0}".format(updateEditingTypeID_SQL), "Error: Updating Editing Type ID...")
     return
    # now load 'edit' page with this new TypeID
    load_MRA_Template_ForEditing(TypeIDtoUse=duplicatedID, originalItemID=origItemID)
  return

# # # #  END OF:  Matter Risk Assessment Templates   # # # #


def Publish_MRA(s, event):
  #! New function 21st August 2025 - This will 'publish' the current 'editing' NMRA (overwriting CaseTypeDefaults)
  #! NB: might make sense to have some row colouring to indicate if a Question has no answers?? (would mean we need a hidden col in DG to store current count)

  # ought to do some checks before allowing 'publish', as we don't want to publish:
  # - template with no questions
  # - Questions with no answers (every Q should have at least one answer)

  myTypeID = tb_ThisMRAid.Text
  oldTypeID = tb_CopyOfMRAid.Text
  if myTypeID is None or myTypeID == '':
    MessageBox.Show("There is no TypeID set for the current MRA template, cannot publish!", "Error: Publishing Matter Risk Assessment...")
    return
  
  # first get count of current items (if no questions - we can't publish)
  countOfQs = _tikitResolver.Resolve("[SQL: SELECT COUNT(ID) FROM Usr_MRA_TemplateQs WHERE TypeID = {0}]".format(myTypeID))
  if int(countOfQs) == 0:
    MessageBox.Show("No questions exist for the selected MRA, cannot publish!", "Error: Publishing Matter Risk Assessment...")
    return
  
  # next, any Questions without any Answers
  tmpQDetails = {"QuestionText": "", "DisplayOrder": 0}
  QsMissingAnswers = []
  checkAnswersSQL = """SELECT TQ.QuestionText, TQ.DisplayOrder
                        FROM Usr_MRA_TemplateQs TQ
                        WHERE TQ.TypeID = {typeID}
                          AND NOT EXISTS (SELECT 1 FROM Usr_MRA_TemplateAs TA WHERE TA.QuestionID = TQ.QuestionID)
                        ORDER BY TQ.DisplayOrder;""".format(typeID=myTypeID)
  
  # now need to run this SQL in the old way so that we can get more data back
  _tikitDbAccess.Open(checkAnswersSQL)
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        tmpItem = tmpQDetails.copy()
        tmpItem['QuestionText'] = '' if dr.IsDBNull(0) else dr.GetString(0)
        tmpItem['DisplayOrder'] = 0 if dr.IsDBNull(1) else dr.GetValue(1)
        QsMissingAnswers.append(tmpItem)
    dr.Close()
  _tikitDbAccess.Close()

  if len(QsMissingAnswers) > 0:
    tmpMsg = "This NMRA cannot be published because the following questions have no answers defined:\n"
    for x in QsMissingAnswers:
      tmpMsg += "- {0} (Display Order: {1})\n".format(x['QuestionText'], x['DisplayOrder'])

    MessageBox.Show(tmpMsg, "Error: Publishing Matter Risk Assessment...")
    return

  # otherwise, if we get this far, we've checked that there are Questions and Answers associated to those Questions, so we can now 'Publish'
  
  # ask for confirmation (as could've clicked button in error)
  confirmPublish = MessageBox.Show("Are you sure you want to publish this Matter Risk Assessment template?\n\nThis will overwrite the current default NMRA used when creating new matters.", "Please Confirm: Publishing Matter Risk Assessment...", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
  if confirmPublish == DialogResult.No:
    return
  else:
    # here we just need to:
    # 1) overwrite 'CaseTypeDefaults' table to set this NEW TypeID
    # 2) set the 'Hidden' flag for THIS new template to 'N' (as it was set to 'Y' when we copied for editing)
    # 3) set the 'Hidden' flag and 'expiry date' of the old template (the one we're replacing), and ideally remove the 'EditingTypeID' value
    #    (was wondering if we should keep this value, just for note of what template replaces current, so we have a lineage of progress?)
    #    (if we did want this, it may be more desirable to have ANOTHER field for this purpose - eg: 'ReplacedByTypeID' or similar)
    #    - or perhaps we could just append text into the 'InternalNote' to this effect (eg: '** This templates is replaced by TypeID: xx from {date} **')
    # ALSO: copy Score Matrix (unless we already do that at the beginning?)

    # should probably have a count, and only update if there is anything to replace
    countSQL = "SELECT COUNT(ID) FROM Usr_MRA_CaseType_Defaults WHERE TypeName = 'Matter Risk Assessment' AND TemplateID = {oldID}".format(oldID=oldTypeID)
    countCaseTypesToUpdate = runSQL(codeToRun=countSQL, showError=False, returnType='Int')

    if int(countCaseTypesToUpdate) > 0:
      # update CaseTypeDefaults for NMRA
      updateSQL = """[SQL: UPDATE Usr_MRA_CaseType_Defaults SET TemplateID = {newID}
                    WHERE TypeName = 'Matter Risk Assessment' AND TemplateID = {oldID}]""".format(newID=myTypeID, oldID=oldTypeID)
    
      try:
        _tikitResolver.Resolve(updateSQL)
        canContinue = True
      except:
        MessageBox.Show("There was an error overwriting the NMRA for all CaseTypes.\n\nUsing SQL:\n{0}".format(updateSQL), "Error: Publish NMRA")
        canContinue = False
    else:
      MessageBox.Show("There don't appear to be any CaseTypes currently using the old Template.\n\nPlease remember to assign this new Template to the appropriate CaseTypes", "Overwrite CaseTypeDefaults with new Template ID...")
      canContinue = True

    if canContinue == True:
      # finally we update the main TemplateTypes table

      # firstly the new item - remove '[Editing]', from name, set to NOT hidden and new 'EffectiveDate'
      updateNewSQL = """UPDATE Usr_MRA_TemplateTypes SET Hidden = 'N', EffectiveFrom = GETDATE(), TypeName = REPLACE(TypeName, ' [Editing]', ''),
                        InternalNote = CONCAT(InternalNote, ' ** This NMRA replaces TypeID: {oldID} - effective: ', CONVERT(varchar(12), GETDATE(), 103), '**')
                        WHERE TypeID = {newID}""".format(oldID=oldTypeID, newID=myTypeID)

      runSQL(codeToRun=updateNewSQL, showError=True, 
             errorMsgText="There was an error amending the details of the new Template.\n\nUsing SQL:\n{0}".format(updateNewSQL),
             errorMsgTitle="Error: Publish NMRA (updating template header details)")

      # lastly, the old item - set to hidden, and set expiry date (effective to) to todays' date
      updateOldSQL = """UPDATE Usr_MRA_TemplateTypes SET Hidden = 'Y', EffectiveTo = GETDATE(), 
                        InternalNote = CONCAT(InternalNote, ' ** This NMRA was replaced by TypeID: {newID} - effective: ', CONVERT(varchar(12), GETDATE(), 103), '**')
                        WHERE TypeID = {oldID}""".format(oldID=oldTypeID, newID=myTypeID)

      runSQL(codeToRun=updateOldSQL, showError=True,
             errorMsgText="There was an error amending the details of the old/replaced Template.\n\nUsing SQL:\n{0}".format(updateOldSQL), 
             errorMsgTitle="Error: Publish NMRA (updating old template header details)")

      MessageBox.Show("This NMRA has now been 'published'.\n\nYou'll be taken back to the main page now - please make sure score thresholds are set appropriately", "Publish NMRA - Complete...")
      
      # refresh the List of MRA templates on main page, and select the item with that new ID
      refresh_MRA_Templates(s, event)

      dg_MRA_Templates.Focus()
      tCount = -1
      for x in dg_MRA_Templates.Items:
        tCount += 1
        if str(x['Code']) == str(myTypeID):
          dg_MRA_Templates.SelectedIndex = tCount
          dg_MRA_Templates.ScrollIntoView(dg_MRA_Templates.Items[tCount])
          break


      ti_MRA_Questions.Visibility = Visibility.Collapsed
      ti_MRA_Overview.Visibility = Visibility.Visible
      ti_MRA_Overview.IsSelected = True
  return

# # # #   S C O R E   T H R E S H O L D S   # # # #
 
def btn_SaveScoreThresholds_Click(s, event):
  # This function does the actual SAVING of the score thresholds
  MRA_ID = dg_MRA_Templates.SelectedItem['Code']
  
  # get current values for current MRA  
  lowTo_SQL = "[SQL: UPDATE Usr_MRA_ScoreMatrix SET Score_To = {0} WHERE TypeID = {1} AND LMH_ID = 1]".format(tb_SM_Low_To.Text, MRA_ID)
  medFrom_SQL = "[SQL: UPDATE Usr_MRA_ScoreMatrix SET Score_From = {0} WHERE TypeID = {1} AND LMH_ID = 2]".format(lbl_SM_Med_From.Content, MRA_ID)
  medTo_SQL = "[SQL: UPDATE Usr_MRA_ScoreMatrix SET Score_To = {0} WHERE TypeID = {1} AND LMH_ID = 2]".format(tb_SM_Med_To.Text, MRA_ID)
  highFrom_SQL = "[SQL: UPDATE Usr_MRA_ScoreMatrix SET Score_From = {0} WHERE TypeID = {1} AND LMH_ID = 3]".format(lbl_SM_High_From.Content, MRA_ID)
  highTo_SQL = "[SQL: UPDATE Usr_MRA_ScoreMatrix SET Score_To = {0} WHERE TypeID = {1} AND LMH_ID = 3]".format(lbl_SM_High_To.Content, MRA_ID)
  
  errCount = 0
  tmpMsg = "There was an error updating the following Score Matrix values:\n"
  
  try: 
    _tikitResolver.Resolve(lowTo_SQL)
  except:
    errCount += 1
    tmpMsg += "- 'Low To', using SQL:\n{0}\n\n".format(lowTo_SQL)
  
  try:
    _tikitResolver.Resolve(medFrom_SQL)
  except:
    errCount += 1
    tmpMsg += "- 'Medium From', using SQL:\n{0}\n\n".format(medFrom_SQL)
    
  try:
    _tikitResolver.Resolve(medTo_SQL)
  except:
    errCount += 1
    tmpMsg += "- 'Medium To', using SQL:\n{0}\n\n".format(medTo_SQL)
    
  try:
    _tikitResolver.Resolve(highFrom_SQL)
  except:
    errCount += 1
    tmpMsg += "- 'High From', using SQL:\n{0}\n\n".format(highFrom_SQL)
    
  try:
    _tikitResolver.Resolve(highTo_SQL)  
  except:
    errCount += 1
    tmpMsg += "- 'High To', using SQL:\n{0}\n\n".format(highTo_SQL) 
  
  if errCount == 0:
    MessageBox.Show("Successfully updated the Score Matrix", "Updating Score Matrix values...")  
  else:
    MessageBox.Show(tmpMsg, "Error: Updating Score Matrix values...")  
  return


def ST_Low_SliderChanged(s, event):
  # This function just updates the 'Mid' minimum value accordingly
  lowValue = int(sld_Low_To.Value)
  lbl_SM_Med_From.Content = str(lowValue + 1)
  sld_Med_To.Minimum = float(lowValue + 2)
  return

def ST_Med_SliderChanged(s, event):
  # This function updates the 'High' from value accordingly
  medValue = int(sld_Med_To.Value)
  lbl_SM_High_From.Content = str(medValue + 1)
  return
  

def setup_ST_Sliders(s, event):
  if int(dg_MRA_Templates.SelectedItem['QCount']) == 0:
    tb_ST_NoQs.Visibility = Visibility.Visible
    grd_ScoreThresholds.Visibility = Visibility.Hidden
    btn_SaveScoreThresholds.Visibility = Visibility.Hidden
    return
  else:
    tb_ST_NoQs.Visibility = Visibility.Hidden
    grd_ScoreThresholds.Visibility = Visibility.Visible  
    btn_SaveScoreThresholds.Visibility = Visibility.Visible  
  
  if int(lbl_Sel_MRA_ID.Content) > 0:
    # get 'MAX' score
    max_SQL = """[SQL: SELECT SUM(Highest_Answer) FROM 
                  (SELECT 'DO' = TQs.DisplayOrder, 'QT' = TQs.QuestionText, 'Highest_Answer' = (SELECT MAX(TAs.Score) FROM Usr_MRA_TemplateAs TAs WHERE TQs.AnswerList = TAs.GroupName) 
                  FROM Usr_MRA_TemplateQs TQs 
                  WHERE TQs.TypeID = {0}) as tmpT]""".format(lbl_Sel_MRA_ID.Content)
    
    try:
      maxValue = _tikitResolver.Resolve(max_SQL)
      canContinue = True
    except:
      canContinue = False
      MessageBox.Show("Couldn't get the 'Maximum' value of question answers, using SQL:\n" + str(max_SQL), "Error: Setting up the Score Threshold sliders...")
      return
    
    if canContinue == True:
      sld_Low_To.Maximum = float(int(maxValue))
      sld_Med_To.Maximum = float((int(maxValue) - 2))
      lbl_SM_High_To.Content = str(maxValue)
  return


def addNew_ScoreThreshold(newTypeID):
  # This function will create a new 'Score Thresholds' table for newly added Matter Risk Assessment
  
  if int(newTypeID) > 0:
    ins_SQL = """[SQL: INSERT INTO Usr_MRA_ScoreMatrix (TypeID, LMH_ID, Score_From, Score_To) 
                  SELECT {0}, LMH_ID, SF, ST FROM 
                  (SELECT 'LMH_ID' = 1, 'SF' = 0, 'ST' = 10 UNION ALL 
                  SELECT 'LMH_ID' = 2, 'SF' = 11, 'ST' = 20 UNION ALL 
                  SELECT 'LMH_ID' = 3, 'SF' = 21, 'ST' = 30) as tmpT]""".format(newTypeID)
    
    try:
      _tikitResolver.Resolve(ins_SQL)
      #refresh_MRA_LMH_ScoreThresholds_New(s, event)
    except:
      MessageBox.Show("There was an error adding new Score Thresholds, using SQL:\n" + str(ins_SQL), "Error: Adding New Score Thresholds...")
  return
  

def duplicate_ScoreThresholds(TypeID_toCopy, newTypeID):
  # This function will duplicate the passed 'TypeID_toCopy' and will give new name (ID)

  ins_SQL = """[SQL: INSERT INTO Usr_MRA_ScoreMatrix (TypeID, LMH_ID, Score_From, Score_To) 
                SELECT {0}, LMH_ID, Score_From, Score_To FROM Usr_MRA_ScoreMatrix WHERE TypeID = {1}]""".format(newTypeID, TypeID_toCopy)

  try:
    _tikitResolver.Resolve(ins_SQL)
    #refresh_MRA_LMH_ScoreThresholds_New(s, event)
  except:
    MessageBox.Show("There was an error duplicating Score Thresholds, using SQL:\n" + str(ins_SQL), "Error: Adding New Score Thresholds...")
  return
 

def refresh_MRA_LMH_ScoreThresholds_New(s, event):
  setup_ST_Sliders(s, event)

  if dg_MRA_Templates.SelectedIndex == -1:
    return
  
  MRA_ID = dg_MRA_Templates.SelectedItem['Code']
  
  # get current values for current MRA  
  lowTo_SQL = "[SQL: SELECT TOP 1 Score_To FROM Usr_MRA_ScoreMatrix WHERE TypeID = {0} AND LMH_ID = 1]".format(MRA_ID)
  medFrom_SQL = "[SQL: SELECT TOP 1 Score_From FROM Usr_MRA_ScoreMatrix WHERE TypeID = {0} AND LMH_ID = 2]".format(MRA_ID)
  medTo_SQL = "[SQL: SELECT TOP 1 Score_To FROM Usr_MRA_ScoreMatrix WHERE TypeID = {0} AND LMH_ID = 2]".format(MRA_ID)
  highFrom_SQL = "[SQL: SELECT TOP 1 Score_From FROM Usr_MRA_ScoreMatrix WHERE TypeID = {0} AND LMH_ID = 3]".format(MRA_ID)
  highTo_SQL = "[SQL: SELECT TOP 1 Score_To FROM Usr_MRA_ScoreMatrix WHERE TypeID = {0} AND LMH_ID = 3]".format(MRA_ID)
  
  tb_SM_Low_To.Text = _tikitResolver.Resolve(lowTo_SQL)
  lbl_SM_Med_From.Content = _tikitResolver.Resolve(medFrom_SQL)
  tb_SM_Med_To.Text = _tikitResolver.Resolve(medTo_SQL)
  lbl_SM_High_From.Content = _tikitResolver.Resolve(highFrom_SQL)
  lbl_SM_High_To.Content = _tikitResolver.Resolve(highTo_SQL)

  return
  
def scoreMatrix_setMax(s, event):
  if int(lbl_Sel_MRA_ID.Content) > 0:
    # get 'MAX' score
    max_SQL = """[SQL: SELECT SUM(Highest_Answer) FROM 
                  (SELECT 'DO' = TQs.DisplayOrder, 'QT' = TQs.QuestionText, 'Highest_Answer' = (SELECT MAX(TAs.Score) FROM Usr_MRA_TemplateAs TAs WHERE TAs.QuestionID = TQs.QuestionID) 
                  FROM Usr_MRA_TemplateQs TQs 
                  WHERE TQs.TypeID = {0}) as tmpT]""".format(lbl_Sel_MRA_ID.Content)
    
    try:
      maxValue = _tikitResolver.Resolve(max_SQL)
      canContinue = True
    except:
      canContinue = False
      MessageBox.Show("There was an error getting the max score, using SQL:\n" + max_SQL, "Error: Getting Max Score...")
    
    if canContinue == True:
      sld_Low_To.Maximum = float(int(maxValue))
      sld_Med_To.Maximum = float((int(maxValue) - 2))
      lbl_SM_High_To.Content = str(maxValue)
  return

# # # #   END OF:   S C O R E   T H R E S H O L D S   # # # #

# # # #   D E P A R T M E N T   D E F A U L T S   # # # #

class MRA_Template_CTG(object):
  def __init__(self, myTicked, myTemplateName, myTemplateID):
     if myTicked == 'True':
       self.Ticked = True
     else:
      self.Ticked = False
     
     self.TName = myTemplateName
     self.TID = myTemplateID
     return
     
  def __getitem__(self, index):
    if index == 'Ticked':
      return self.Ticked
    elif index == 'TName':
      return self.TName
    elif index == 'TID':
      return self.TID
    else:
      return ''
      
def refresh_MRA_Department_Defaults(s, event):
  # This function will populate the DEPARTMENT Defaults datagrid
  # SQL to populate datagrid
  getTableSQL = "SELECT 'Ticked' = 'False', TypeName, TypeID FROM Usr_MRA_TemplateTypes WHERE Is_MRA = 'Y' AND Hidden = 'N' ORDER BY TypeName"
  
  tmpItem = []
  _tikitDbAccess.Open(getTableSQL)
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          tmpTick = 'False' if dr.IsDBNull(0) else dr.GetString(0)
          tmpTName = '' if dr.IsDBNull(1) else dr.GetString(1)
          tmpTid = 0 if dr.IsDBNull(2) else dr.GetValue(2)
        
          tmpItem.append(MRA_Template_CTG(tmpTick, tmpTName, tmpTid))

    dr.Close()
  
  #close db connection
  _tikitDbAccess.Close()
  dg_MRA_Templates_CTD.ItemsSource = tmpItem
  return


def MRA_Department_Defaults_SelectionChanged(s, event):
  # This function will cause the Case Type list to the right to update to show selected department
  if dg_MRA_Templates_CTD.SelectedIndex == -1:
    lbl_SelectedDeptID.Content = ''
    lbl_SelectedDeptName.Content = ''
    #btn_Save_MRA_TemplateToUseForDept.IsEnabled = False
    return

  btn_Save_MRA_TemplateToUseForDept.IsEnabled = True
  #refresh_MRA_CaseType_Defaults(s, event)
  return
  

def MRA_Save_Default_For_Department(s, event):
  # This function will save the current selection to all the Case Types within the selected Department 
  # NB: will check to see if anything has been assigned currently and prompts if OK to overwrite
  
  deptRef = lbl_SelectedDeptID.Content
  
  countExisting_SQL = "SELECT COUNT(ID) FROM Usr_MRA_CaseType_Defaults WHERE CaseTypeID IN (SELECT Code FROM CaseTypes WHERE CaseTypeGroupRef = {0}) AND TypeName = 'Matter Risk Assessment' AND TemplateID > -1".format(deptRef)
  countExisting = runSQL(countExisting_SQL, False, '', '')
  
  if int(countExisting) > 0:
    myResult = MessageBox.Show("There appear to defaults already set against some Case Types in this Department, are you sure you wish to overwrite changes?", "Overwrite existing defaults...", MessageBoxButtons.YesNo)
    if myResult == DialogResult.No: 
      return
  
  # Firstly, delete all selections for current items
  del_SQL = "DELETE FROM Usr_MRA_CaseType_Defaults WHERE TypeName = 'Matter Risk Assessment' AND CaseTypeID IN (SELECT Code FROM CaseTypes WHERE CaseTypeGroupRef = {0})".format(deptRef)
  runSQL(del_SQL, True, "There was an error deleting the old template types associated to Case Types in selected department", "Error: Update defaults for all Case Types in Dept...")
  
  # now iterate over datagrid and add new items  
  for dgItem in dg_MRA_Templates_CTD.Items:
    if dgItem.Ticked == True:
      ins_SQL = "INSERT INTO Usr_MRA_CaseType_Defaults (CaseTypeID, TemplateID, TypeName) "
      ins_SQL += "SELECT Code, {0}, 'Matter Risk Assessment' FROM CaseTypes WHERE CaseTypeGroupRef = {1} AND Description NOT LIKE '%Project%'".format(dgItem.TID, deptRef)
      runSQL(ins_SQL, True, "There was an error adding the new template associations for the selected department", "Error: Update defaults for all Case Types in Dept...")

  # refresh Case Types Defaults list...
  refresh_MRA_CaseType_Defaults(s, event)
  refresh_MRA_Templates(s, event)
  return
  

# # # #   END OF:   D E P A R T M E N T   D E F A U L T S   # # # #

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
      
def refresh_MRA_CaseType_Defaults(s, event):
  # This function will populate the 'Case Types' datagrid (for selecting which Matter Risk Assessment template to be used)
  #! New 04/09/2025: Added SQL to add any 'missing' caseTypes that we don't have in our 'CaseTypeDefaults' table
  add_Missing_CaseTypeDefaults(forWhat='Matter Risk Assessment')

  getTableSQL = """SELECT '0-RowID' = ISNULL(STRING_AGG(CTD.ID, ';'), ''), 
                          '1-CaseType Name' = CT.Description, 
                          '2-MRA TemplateID' = ISNULL(STRING_AGG(CTD.TemplateID, '|'), ''), 
                          '3-MRA Template Name' = ISNULL(STRING_AGG(TT.TypeName, '; '), ''), 
                          '4-CaseType ID' = CT.Code, 
                          '5-CaseTypeGroup Name' = CTG.Name, 
                          '6-CTG ID' = CT.CaseTypeGroupRef 
                  FROM CaseTypes CT 
                      LEFT OUTER JOIN CaseTypeGroups CTG ON CT.CaseTypeGroupRef = CTG.ID 
                      LEFT OUTER JOIN Usr_MRA_CaseType_Defaults CTD ON CT.Code = CTD.CaseTypeID AND CTD.TypeName = 'Matter Risk Assessment' 
                      LEFT OUTER JOIN Usr_MRA_TemplateTypes TT ON CTD.TemplateID = TT.TypeID 
                  GROUP BY CT.Description, CT.Code, CTG.Name, CT.CaseTypeGroupRef ORDER BY CT.Description """

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
          
          tmpItem.append(caseType_Defaults(tmpCaseTypeName, tmpTemplateName, tmpRowID, tmpTemplateID, 
                                           tmpCaseTypeID, tmpCaseTypeGroupName, tmpCTGID))

    dr.Close()
  
  #close db connection
  _tikitDbAccess.Close()
  
  # added the following 19th June 2024 - this will display list grouped by 'Department' (CaseTypeGroup) name
  # note: added ', CollectionView, ListCollectionView, PropertyGroupDescription' to 'from System.Windows.Data import Binding ' (line 20)
  tmpC = ListCollectionView(tmpItem)
  tmpC.GroupDescriptions.Add(PropertyGroupDescription("mraT_CTGName"))

  dg_MRA_CaseTypes_MRATemplate.ItemsSource = tmpC     #tmpItem
  return


def MRA_CaseType_Defaults_SelectionChanged(s, event):
  # This function populates the controls beneath the data grid to allow for updates to be made
  if dg_MRA_CaseTypes_MRATemplate.SelectedIndex == -1:
    lbl_SelectedCaseType.Content = ''
    lbl_SelectedCaseTypeID.Content = ''
    lbl_SelectedDeptID.Content = ''
    lbl_SelectedDeptName.Content = ''
    btn_Save_MRA_TemplateToUseForCaseType.IsEnabled = False
    return

  btn_Save_MRA_TemplateToUseForCaseType.IsEnabled = True
  lbl_SelectedCaseType.Content = dg_MRA_CaseTypes_MRATemplate.SelectedItem['CTName']
  lbl_SelectedCaseTypeID.Content = dg_MRA_CaseTypes_MRATemplate.SelectedItem['CTID']
  lbl_SelectedDeptID.Content = dg_MRA_CaseTypes_MRATemplate.SelectedItem['CTGid']
  lbl_SelectedDeptName.Content = dg_MRA_CaseTypes_MRATemplate.SelectedItem['CTGName']
  tempIDs = str(dg_MRA_CaseTypes_MRATemplate.SelectedItem['TemplateID'])
  #MessageBox.Show("TemplateIDs: " + str(tempIDs))
  
  dgItems = []
  if "|" in tempIDs:
    tIDs = []
    tIDs = tempIDs.split('|')
  
    for xRow in dg_MRA_Templates_CTD.Items:
      aTicked = 'False'
      aName = xRow.TName
      aID = xRow.TID
  
      if len(tIDs) > 0:
        for idItem in tIDs:
          #MessageBox.Show("aID = " + str(aID) + "\nidItem = " + str(idItem), "Debug Testing")
          if str(aID) == str(idItem):
            aTicked = 'True'
            break
    
      #MessageBox.Show("aTicked: " + aTicked + "\naName: " + aName + "\naID: " + str(aID), "Debug Testing")
      dgItems.append(MRA_Template_CTG(aTicked, aName, aID))
  else:
    tID = tempIDs
    for xRow in dg_MRA_Templates_CTD.Items:
      aTicked = 'False'
      aName = xRow.TName
      aID = xRow.TID
  
      #MessageBox.Show("aID = " + str(aID) + "\ntID = " + str(tID), "Debug Testing")
      if str(aID) == str(tID):
        aTicked = 'True'
    
      #MessageBox.Show("aTicked: " + aTicked + "\naName: " + aName + "\naID: " + str(aID), "Debug Testing")
      dgItems.append(MRA_Template_CTG(aTicked, aName, aID))    
    
  dg_MRA_Templates_CTD.ItemsSource = dgItems
  
  return
  
  
def MRA_Save_Default_For_CaseType(s, event):
  # This function saves the selected template to the Case Type defaults table
  
  CTid = lbl_SelectedCaseTypeID.Content
  CTName = lbl_SelectedCaseType.Content
  
  # Firstly, delete all selections for current items
  del_SQL = "DELETE FROM Usr_MRA_CaseType_Defaults WHERE TypeName = 'Matter Risk Assessment' AND CaseTypeID = ".format(CTid)
  runSQL(del_SQL, True, "There was an error deleting the old template types associated to the selected Case Type ({0} - {1})".format(CTName, CTid), "Error: Update defaults for selected Case Type...")
  
  # now iterate over datagrid and add new items  
  for dgItem in dg_MRA_Templates_CTD.Items:
    if dgItem.Ticked == True:
      ins_SQL = """INSERT INTO Usr_MRA_CaseType_Defaults (CaseTypeID, TemplateID, TypeName) 
                    VALUES({0}, {1}, 'Matter Risk Assessment')""".format(CTid, dgItem.TID)
      runSQL(ins_SQL, True, "There was an error adding the template '{0}' for the current Case Type '{1}' ({2})".format(dgItem.TName, CTName, CTid), "Error: Add Template default for Case Type...")
    
  refresh_MRA_CaseType_Defaults(s, event)  
  refresh_MRA_Templates(s, event)
  return

#def tog_ExpandOrContract(s, event):
## This was supposed to act as an 'expand all' / 'contract all' button to show/hide all grouped rows in datagrid
## However, we get error "'none type' object has no attribute 'IsExpanded'", so will need more digging to get this working
## But parking for now (maybe on an update I can re-address this).
#  if tog_ExpandContract.IsChecked == True:
#    exp.IsExpanded = False
#    tog_ExpandContract.Content = 'Expand All'
#  else:
#    exp.IsExpanded = True
#    tog_ExpandContract.Content = 'Contract All'
#  
#  return

# # # #   END OF:   C A S E   T Y P E   D E F A U L T S   # # # #


# # # #   Q U E S T I O N S   # # # #

class MRA_Questions(object):
  def __init__(self, myID, myDO, myText, mySection, myAList, mySectionID = 0, myRowID = 0):
    self.mraQ_ID = myID
    self.mraQ_DO = myDO
    self.mraQ_Text = myText
    self.mraQ_Section = mySection
    self.mraQ_AList = myAList
    self.mraQ_GroupID = mySectionID
    self.mraQ_RowID = myRowID
    return
    
  def __getitem__(self, index):
    if index == 'ID':
      return self.mraQ_ID
    elif index == 'Order':
      return self.mraQ_DO
    elif index == 'QText':
      return self.mraQ_Text
    elif index == 'Section':
      return self.mraQ_Section
    elif index == 'AList':
      return self.mraQ_AList
    elif index == 'SectionID':
      return self.mraQ_GroupID
    elif index == 'RowID':
      return self.mraQ_RowID
    else:
      return ''

def refresh_MRA_Questions(for_MRA_ID=0):
  # This function will populate the QUESTIONS datagrid (dg_MRA_Questions)
  
  if for_MRA_ID == 0:
    tb_NoQuestions_MRA.Visibility = Visibility.Visible
    dg_MRA_Questions.Visibility = Visibility.Hidden
    return

  getTableSQL = """SELECT MRATQ.DisplayOrder, 
                      MRATQ.QuestionText, 
                      'Group' = (SELECT Name FROM Usr_MRA_QGroups QG WHERE QG.ID = MRATQ.QGroupID), 
                      'Num of Answers' = (SELECT COUNT(ID) FROM Usr_MRA_TemplateAs WHERE QuestionID = MRATQ.QuestionID), 
                      MRATQ.QuestionID, MRATQ.QGroupID, MRATQ.ID 
                   FROM Usr_MRA_TemplateQs MRATQ 
                   WHERE MRATQ.TypeID = {0} 
                   ORDER BY MRATQ.QGroupID, MRATQ.DisplayOrder""".format(for_MRA_ID)    #lbl_EditRiskAssessment_ID.Content)
  
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
          tmpNumAnswers = '' if dr.IsDBNull(3) else dr.GetValue(3)
          tmpQID = 0 if dr.IsDBNull(4) else dr.GetValue(4)
          tmpTabGroupID = 0 if dr.IsDBNull(5) else dr.GetValue(5)
          tmpRowID = 0 if dr.IsDBNull(6) else dr.GetValue(6)
          
          tmpItem.append(MRA_Questions(myID=tmpQID, myDO=tmpDO, myText=tmpQText, 
                                       mySection=tmpTabGroup, myAList=tmpNumAnswers, 
                                       mySectionID=tmpTabGroupID, myRowID=tmpRowID))

    dr.Close()
  
  #close db connection
  _tikitDbAccess.Close()

  # if no grouping, use the code below
  #dg_MRA_Questions.ItemsSource = tmpItem
  #! updated 26/08/2025 to show groups instead of flat list
  tmpC = ListCollectionView(tmpItem)
  tmpC.GroupDescriptions.Add(PropertyGroupDescription("mraQ_Section"))
  dg_MRA_Questions.ItemsSource = tmpC
  
  if dg_MRA_Questions.Items.Count > 0:
    tb_NoQuestions_MRA.Visibility = Visibility.Hidden
    dg_MRA_Questions.Visibility = Visibility.Visible
  else:
    tb_NoQuestions_MRA.Visibility = Visibility.Visible
    dg_MRA_Questions.Visibility = Visibility.Hidden
  return


def MRA_Questions_SelectionChanged(s, event):
  # Occurs when user clicks on a Question row on the 'Editing Questions' of a MRA
  
  if dg_MRA_Questions.SelectedIndex > -1:
    lbl_QID.Content = dg_MRA_Questions.SelectedItem['ID']
    txt_QuestionText.Text = dg_MRA_Questions.SelectedItem['QText']
    
    # iterate over items in Question Group list and select matching item from list
    pCount = -1
    for xRow in cbo_QuestionGroup.Items:
      pCount += 1
      if xRow.Name == dg_MRA_Questions.SelectedItem['Section']:
        cbo_QuestionGroup.SelectedIndex = pCount
        break
    
    # NEW Need to load Answers list for selected Q
    populate_AnswersPreview(s, event)
    
  else:
    txt_QuestionText.Text = ''
    cbo_QuestionGroup.SelectedIndex = -1
    cbo_QuestionAnswerList.SelectedIndex = -1
    
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


def SaveChanges_MRA_Question(s, event):
  # This function will save the changes made in the 'edit' area back to the database (and refreshes the data grid)
  #! NOTE: 02/09/2025 - THIS has been amended so that we're taking 'Group' into account...
  #! Eg: if user changes Group, we update the DisplayOrder so that new item is at the end of the new group
  #!     and also need to re-order the old group to close any gaps.
  #! So first, we'll need to x-ref the selected 'Group' to see if different from what's currently in the DataGrid.
  #! With this, we should be able to find out 'Previous' GroupID that we need to re-sort/order

  # if something in'Q ID' label, we can proceed
  if len(str(lbl_QID.Content)) > 0:
    # escape single quotes in question text
    tmpQ = str(txt_QuestionText.Text)
    tmpQ = tmpQ.replace("'", "''")

    # get current 'Group' from drop-down
    newQGroup = cbo_QuestionGroup.SelectedItem['Code']
    # get previous 'Group' from the datagrid
    prevQGroup = dg_MRA_Questions.SelectedItem['SectionID']
    # we also need to make sure we set the DisplayOrder to be at the end of the new group
    if newQGroup != prevQGroup:
      newDO = _tikitResolver.Resolve("[SQL: SELECT ISNULL(MAX(DisplayOrder) + 1, 1) FROM Usr_MRA_TemplateQs WHERE TypeID = {typeID} AND QGroupID = {qGroupID}]".format(typeID=tb_ThisMRAid.Text, qGroupID=newQGroup))
      update_SQL = """[SQL: UPDATE Usr_MRA_TemplateQs SET QuestionText = '{qText}', QGroupID = {qGroupID}, DisplayOrder = {qDO} 
                          WHERE QuestionID = {qID} AND TypeID = {typeID}]""".format(qText=tmpQ, qGroupID=newQGroup, qID=lbl_QID.Content, typeID=tb_ThisMRAid.Text, qDO=newDO)   

    else:
      # FIRST: update details so that item is removed from previous group
      update_SQL = """[SQL: UPDATE Usr_MRA_TemplateQs SET QuestionText = '{qText}', QGroupID = {qGroupID} 
                          WHERE QuestionID = {qID} AND TypeID = {typeID}]""".format(qText=tmpQ, qGroupID=newQGroup, qID=lbl_QID.Content, typeID=tb_ThisMRAid.Text)
    
    try:
      _tikitResolver.Resolve(update_SQL)
    except:
      MessageBox.Show("There was an error saving Question details, using SQL:\n" + str(update_SQL), "Error: Saving Question details (Matter Risk Assessment)...")
      return
    
    # if Group was changed, we need to re-order the 'old' group to close any gaps
    if newQGroup != prevQGroup:
      # first, re-order the 'old' group to close any gaps
      reOrderOldGroup_SQL = """WITH CTE AS 
                                (SELECT DisplayOrder, 
                                        NewDO = ROW_NUMBER() OVER (ORDER BY DisplayOrder) 
                                 FROM Usr_MRA_TemplateQs 
                                 WHERE TypeID = {typeID} AND QGroupID = {qGroupID}) 
                                UPDATE CTE SET DisplayOrder = NewDO""".format(typeID=tb_ThisMRAid.Text, qGroupID=prevQGroup)
      
      # execute the SQL
      runSQL(codeToRun=reOrderOldGroup_SQL, showError=True, 
             errorMsgText="There was an error re-ordering the previous Question Group to close any gaps, using SQL:\n" + str(reOrderOldGroup_SQL), 
             errorMsgTitle="Error: Re-ordering previous Question Group...", 
             useAltResolver=True)
    
    # refresh list
    refresh_MRA_Questions(for_MRA_ID=tb_ThisMRAid.Text)

  else:
    MessageBox.Show("Nothing to save!", "Save Question details (Matter Risk Assessment)...") 
  return


# Editable Answers on 'Editing Questions' tab - selection changed
def dg_EditMRA_AnswersPreview_SelectionChanged(s, event):
  if dg_EditMRA_AnswersPreview.SelectedIndex > -1:
    lbl_MRA_Answer_Text.Content = dg_EditMRA_AnswersPreview.SelectedItem['AText']
    lbl_MRA_Answer_Score.Content = dg_EditMRA_AnswersPreview.SelectedItem['AScore']
    lbl_MRA_Answer_EmailComment.Content = dg_EditMRA_AnswersPreview.SelectedItem['EmailComment']
  else:
    lbl_MRA_Answer_Text.Content = ''
    lbl_MRA_Answer_Score.Content = ''
    lbl_MRA_Answer_EmailComment.Content = ''
  return


def dg_EditMRA_AnswersPreview_CellEditEnding(s, event):
  # This function commits changes to Editable Answers on 'Editing Questions' tab - cell edit ending
  tmpCol = event.Column
  tmpColName = tmpCol.Header
  
  # Get Initial values updated
  updateSQL = "UPDATE Usr_MRA_TemplateAs SET "
  countOfUpdates = 0
  itemID = dg_EditMRA_AnswersPreview.SelectedItem['ID']
  newName = dg_EditMRA_AnswersPreview.SelectedItem['AText']
  newNameSql = newName.replace("'", "''")
  newScore = dg_EditMRA_AnswersPreview.SelectedItem['AScore']
  newEC = dg_EditMRA_AnswersPreview.SelectedItem['EmailComment']
  newECsql = newEC.replace("'", "''")

  # Conditionally add parts depending on column updated and whether value has changed
  if tmpColName == 'Answer Text':
    if newName != lbl_MRA_Answer_Text.Content:
      updateSQL += "AnswerText = '{0}' ".format(newNameSql) 
      countOfUpdates += 1
    
  elif tmpColName == 'Score':
    if newScore != lbl_MRA_Answer_Score.Content:
      updateSQL += "Score = {0} ".format(newScore) 
      countOfUpdates += 1
    
  elif tmpColName == 'Email Comment':
    if newEC != lbl_MRA_Answer_EmailComment.Content:
      updateSQL += "EmailComment = '{0}' ".format(newECsql) 
      countOfUpdates += 1
    
  # Add WHERE clause
  updateSQL += "WHERE AnswerID = {0}".format(itemID)
    
  # Only run if something was changed
  if countOfUpdates > 0:
    try:
      runSQL(updateSQL, True, "There was an error trying to update Answer list item.", "Error: Updating Answer list item...")
      populate_AnswersPreview(s, event)
    except:
      return
  return


# toolbar buttons (add new, duplicate, move to top, move up, move down, move to bottom, delete)
def dg_EditMRA_AnswersPreview_addNew(s, event):
  # This function will add a new list item to the currently selected group
  # NB: As of Friday 26th April, we're now tying Answers DIRECTLY to questions via 'QuestionID'.
  # Therefore we have ability to copy answers from another question etc
  newDO = _tikitResolver.Resolve("[SQL: SELECT ISNULL(MAX(DisplayOrder) + 1, 1) FROM Usr_MRA_TemplateAs WHERE QuestionID = {0}]".format(lbl_QID.Content))
  nextAnsID = _tikitResolver.Resolve("[SQL: SELECT ISNULL(MAX(AnswerID), 1) + 1 FROM Usr_MRA_TemplateAs]")
  insert_SQL = """INSERT INTO Usr_MRA_TemplateAs (QuestionID, AnswerText, Score, DisplayOrder, AnswerID) 
                  VALUES({0}, '(new)', 0, {1}, {2})""".format(lbl_QID.Content, newDO, nextAnsID)
  
  runSQL(codeToRun=insert_SQL, showError=True, 
         errorMsgText="There was an error adding a new list item", 
         errorMsgTitle="Error: Add New Answer List item...")
  # also update AnswerID to ID for newly added item
  #runSQL("UPDATE Usr_MRA_TemplateAs SET AnswerID = ID WHERE AnswerID IS NULL", False, '', '')
  #! ^ nope - this is a BAD idea (see other notes - don't want to potentially overlap IDs)
  
  # auto select last item...
  populate_AnswersPreview(s, event)
  dg_EditMRA_AnswersPreview.SelectedIndex = (dg_EditMRA_AnswersPreview.Items.Count - 1)
  return
  

def dg_EditMRA_AnswersPreview_duplicate(s, event):
  #

  # This function will duplicate the currently selected list item
  if dg_EditMRA_AnswersPreview.SelectedIndex == -1:
    MessageBox.Show("Nothing selected to duplicate!", "Error: Duplicating Selected Answer item...")
    return
  
  nextAnsID = _tikitResolver.Resolve("[SQL: SELECT ISNULL(MAX(AnswerID), 1) + 1 FROM Usr_MRA_TemplateAs]")
  selectedID = dg_EditMRA_AnswersPreview.SelectedItem['ID']
  insert_SQL = """INSERT INTO Usr_MRA_TemplateAs (GroupName, AnswerText, Score, DisplayOrder, EmailComment, QuestionID, AnswerID) 
                  SELECT TA.GroupName, CONCAT(TA.AnswerText, ' (copy)'), TA.Score, (SELECT MAX(TA1.DisplayOrder) + 1 FROM Usr_MRA_TemplateAs TA1 WHERE TA1.QuestionID = TA.QuestionID), 
                    EmailComment, QuestionID, {0} 
                  FROM Usr_MRA_TemplateAs TA WHERE TA.AnswerID = """.format(selectedID, nextAnsID)
  
  runSQL(codeToRun=insert_SQL, showError=True, 
         errorMsgText="There was an error duplicating the selected list item", 
         errorMsgTitle="Error: Duplicate Selected Answer List item...")
  # also update AnswerID to ID for newly added item
  #runSQL("UPDATE Usr_MRA_TemplateAs SET AnswerID = ID WHERE AnswerID IS NULL", False, '', '')
  #! ^ nope - this is a BAD idea (see other notes - don't want to potentially overlap IDs)
  
  # auto select last item...
  populate_AnswersPreview(s, event)
  dg_EditMRA_AnswersPreview.SelectedIndex = (dg_EditMRA_AnswersPreview.Items.Count - 1)
  return


def dg_EditMRA_AnswersPreview_moveToTop(s, event):
  # This function will move the selected ANSWER to the top row (and all other items down one)

  tmpNewFocusRow = dgItem_MoveToTop(dgControl=dg_EditMRA_AnswersPreview, 
                                    tableToUpdate='Usr_MRA_TemplateAs', 
                                    sqlOrderColName='DisplayOrder', 
                                    dgIDcolName='RowID', 
                                    dgOrderColName='Order', 
                                    sqlOtherCheckCol='QuestionID', 
                                    sqlOtherCheckValue=lbl_QID.Content)
  if tmpNewFocusRow > -1:
    populate_AnswersPreview(s, event)
    dg_EditMRA_AnswersPreview.Focus()
    dg_EditMRA_AnswersPreview.SelectedIndex = tmpNewFocusRow
  return


def dg_EditMRA_AnswersPreview_moveUp(s, event):
  # This function will move the selected ANSWER up one row (and all other items down one)

  tmpNewFocusRow = dgItem_MoveUp(dgControl=dg_EditMRA_AnswersPreview, 
                                 tableToUpdate='Usr_MRA_TemplateAs', 
                                 sqlOrderColName='DisplayOrder', 
                                 dgIDcolName='RowID', 
                                 dgOrderColName='Order', 
                                 sqlOtherCheckCol='QuestionID', 
                                 sqlOtherCheckValue=lbl_QID.Content)
  if tmpNewFocusRow > -1: 
    populate_AnswersPreview(s, event)
    dg_EditMRA_AnswersPreview.Focus()
    dg_EditMRA_AnswersPreview.SelectedIndex = tmpNewFocusRow
  return
  
  
def dg_EditMRA_AnswersPreview_moveDown(s, event):
  # This function will move the selected Answer down one row (and all other items up one)
  
  tmpNewFocusRow = dgItem_MoveDown(dgControl=dg_EditMRA_AnswersPreview, 
                                   tableToUpdate='Usr_MRA_TemplateAs', 
                                   sqlOrderColName='DisplayOrder', 
                                   dgIDcolName='RowID', 
                                   dgOrderColName='Order', 
                                   sqlOtherCheckCol='QuestionID', 
                                   sqlOtherCheckValue=lbl_QID.Content) 
  if tmpNewFocusRow > -1: 
    populate_AnswersPreview(s, event)
    dg_EditMRA_AnswersPreview.Focus()
    dg_EditMRA_AnswersPreview.SelectedIndex = tmpNewFocusRow  
  return
  
  
def dg_EditMRA_AnswersPreview_moveToBottom(s, event):
  # This function will move the selected Answer to the bottom row (and all other items up one)
  
  tmpNewFocusRow = dgItem_MoveToBottom(dgControl=dg_EditMRA_AnswersPreview, 
                                       tableToUpdate='Usr_MRA_TemplateAs', 
                                       sqlOrderColName='DisplayOrder', 
                                       dgIDcolName='RowID', 
                                       dgOrderColName='Order', 
                                       sqlOtherCheckCol='QuestionID', 
                                       sqlOtherCheckValue=lbl_QID.Content)
  if tmpNewFocusRow > -1:
    populate_AnswersPreview(s, event)
    dg_EditMRA_AnswersPreview.Focus()
    dg_EditMRA_AnswersPreview.SelectedIndex = tmpNewFocusRow  
  return  
  
  
def dg_EditMRA_AnswersPreview_deleteSelected(s, event):
  # This function will delete the currently selected list item
  tmpNewFocusRow = dgItem_DeleteSelected(dgControl=dg_EditMRA_AnswersPreview, 
                                         tableToUpdate='Usr_MRA_TemplateAs', 
                                         sqlOrderColName='DisplayOrder', 
                                         dgIDcolName='RowID', 
                                         dgOrderColName='Order', 
                                         dgNameDescColName='AText', 
                                         sqlOtherCheckCol='QuestionID', 
                                         sqlOtherCheckValue=lbl_QID.Content)
  if tmpNewFocusRow > -1:
    populate_AnswersPreview(s, event)
    dg_EditMRA_AnswersPreview.Focus()
    dg_EditMRA_AnswersPreview.SelectedIndex = tmpNewFocusRow  
  return



# Now for the   Q U E S T I O N S   toolbar buttons
   
def AddNew_MRA_Question(s, event):
  # This function will add a new Question row to the Questions datagraid

  mraIDtoUse = tb_ThisMRAid.Text
  #! Linked to XAML control.event: btn_AddNew_Q.Click
  newDO = _tikitResolver.Resolve("[SQL: SELECT ISNULL(MAX(DisplayOrder) + 1, 1) FROM Usr_MRA_TemplateQs WHERE TypeID = {0}]".format(mraIDtoUse))  #lbl_EditRiskAssessment_ID.Content))
  #! New 25/07/2025 - getting the NEXT available 'QuestionID' to use for new items added (we add on initial insert now rather than doing at end of function)
  #! NB: we '+1' here
  nextQID = _tikitResolver.Resolve("[SQL: SELECT MAX(QuestionID) + 1 FROM Usr_MRA_TemplateQs]")
  insert_SQL = """[SQL: INSERT INTO Usr_MRA_TemplateQs (TypeID, DisplayOrder, QuestionText, AnswerList, QGroupID, QuestionID) 
                      VALUES({0}, {1}, '(new_question)', '', null, {2})]""".format(mraIDtoUse, newDO, nextQID)
  
  runSQL(codeToRun=insert_SQL, showError=True, errorMsgText="There was an error adding a new question to this Matter Risk Assessment", errorMsgTitle="Error: Adding new Question to Matter Risk Assessment...")
  ## also update 'QuestionID'
  #runSQL("UPDATE Usr_MRA_TemplateQs SET QuestionID = ID WHERE QuestionID IS NULL", False, '', '')
  #! ^ nope - this is a BAD idea, as we could overlap with existing QuestionIDs, so we now get the next available ID and use that on insert above

  # auto select last item...
  refresh_MRA_Questions(for_MRA_ID=mraIDtoUse)
  dg_MRA_Questions.SelectedIndex = (dg_MRA_Questions.Items.Count - 1)
  txt_QuestionText.Focus()
  return
  
  
def Duplicate_MRA_Question(s, event):
  #! Linked to XAML control.event: btn_CopySelected_Q.Click
  # This function will duplicate the selected Question and select this duped row so that it's available for editing
  if dg_MRA_Questions.SelectedIndex == -1:
    MessageBox.Show("Nothing selected to duplicate!", "Error: Duplicating Selected Question...")
    return
  
  sourceID = dg_MRA_Questions.SelectedItem['ID']
  mra_TypeID = tb_ThisMRAid.Text   #lbl_EditRiskAssessment_ID.Content
  newDO = _tikitResolver.Resolve("[SQL: SELECT ISNULL(MAX(DisplayOrder) + 1, 1) FROM Usr_MRA_TemplateQs WHERE TypeID = {0}]".format(mra_TypeID))
  #! New 28/5/2025 - getting the NEXT available 'QuestionID' to use for new items added (we add on initial insert now rather than doing at end of function)
  nextQID = _tikitResolver.Resolve("[SQL: SELECT MAX(QuestionID) + 1 FROM Usr_MRA_TemplateQs]")

  # copy Question item
  insert_SQL = """[SQL: INSERT INTO Usr_MRA_TemplateQs (TypeID, DisplayOrder, QuestionText, AnswerList, QGroupID, QuestionID) 
                        SELECT TQ.TypeID, {0}, TQ.QuestionText + ' (copy)', TQ.AnswerList, TQ.QGroupID, {1}   
                        FROM Usr_MRA_TemplateQs TQ WHERE TQ.QuestionID = {2}]""".format(newDO, nextQID, sourceID)
  _tikitResolver.Resolve(insert_SQL)
  
  # get the ID as we'll need this for copying the Answers
  #newQID = _tikitResolver.Resolve("[SQL: SELECT ID FROM Usr_MRA_TemplateQs WHERE TypeID = {0} AND DisplayOrder = {1} AND QuestionID = {2}]".format(mra_TypeID, newDO, nextQID))
  #! ^ doesn't look like we're using 'ID' anymore, so we don't need this
  ## update QuestionID
  #runSQL("UPDATE Usr_MRA_TemplateQs SET QuestionID = ID WHERE QuestionID IS NULL", False, '', '')
  #! ^ nope - this is a BAD idea, as we could overlap with existing QuestionIDs, so we now get the next available ID and use that on insert above
  
  # Before we can copy Answer items, we need to get the next available AnswerID and use this in conjunction with 'DisplayOrder' to ensure we don't overlap with existing items
  nextAnsID = _tikitResolver.Resolve("[SQL: SELECT ISNULL(MAX(AnswerID), 1) FROM Usr_MRA_TemplateAs]")
  # then copy Answers items 
  #insert_SQL = """[SQL: INSERT INTO Usr_MRA_TemplateAs (GroupName, AnswerText, Score, DisplayOrder, EmailComment, QuestionID, AnswerID) 
  #                      SELECT GroupName, AnswerText, Score, DisplayOrder, EmailComment, {0}, null 
  #                      FROM Usr_MRA_TemplateAs WHERE QuestionID = {1}]""".format(nextQID, sourceID)
  #_tikitResolver.Resolve(insert_SQL)

  insSQL = """WITH myCopyList AS (
                  SELECT 'GrpName' = GroupName, 'QID' = {currQID}, 'AnsText' = AnswerText, 'Score' = Score, 
                      'EmailComment' = EmailComment, 'RowNum' = ROW_NUMBER() OVER (ORDER BY DisplayOrder, AnswerText),  
                      'DispOrder' = DisplayOrder 
                  FROM Usr_MRA_TemplateAs WHERE QuestionID = {QtoCopyID} 
                )
              INSERT INTO Usr_MRA_TemplateAs (GroupName, QuestionID, AnswerText, Score, EmailComment, DisplayOrder, AnswerID) 
              SELECT GrpName, QID, AnsText, Score, EmailComment, RowNum, {nextAnswerID} + RowNum FROM myCopyList ORDER BY RowNum;
              """.format(currQID=nextQID, QtoCopyID=sourceID, nextAnswerID=nextAnsID)

  runSQL(codeToRun=insSQL, showError=True, 
         errorMsgText="There was an error adding the Answer list...", 
         errorMsgTitle="Error: adding Answer list...", 
         useAltResolver=True)

  # and update AnswerID
  #runSQL("UPDATE Usr_MRA_TemplateAs SET AnswerID = ID WHERE AnswerID IS NULL", False, '', '')
  #! ^ nope - this is a BAD idea, as we could overlap with existing AnswerIDs, so we now get the next available ID and use that on insert above
  
  # auto select last item...
  refresh_MRA_Questions(for_MRA_ID=mra_TypeID)
  dg_MRA_Questions.SelectedIndex = (dg_MRA_Questions.Items.Count - 1)
  txt_QuestionText.Focus()
  return
  
  
def MoveTop_MRA_Question(s, event):
  # This function will move the selected Question to the top row (and all other items down one)
  
  selectedID = int(dg_MRA_Questions.SelectedItem['ID'])
  tmpNewFocusRow = dgItem_MoveToTop(dgControl=dg_MRA_Questions, 
                                    tableToUpdate='Usr_MRA_TemplateQs', 
                                    sqlOrderColName='DisplayOrder', 
                                    dgIDcolName='RowID', 
                                    dgOrderColName='Order', 
                                    sqlOtherCheckCol='TypeID', 
                                    sqlOtherCheckValue=int(tb_ThisMRAid.Text), 
                                    sqlGroupColName='QGroupID',
                                    dgGroupColName='SectionID')  
  if tmpNewFocusRow > -1:
    refresh_MRA_Questions(for_MRA_ID=tb_ThisMRAid.Text)
    dg_MRA_Questions.Focus()
    dg_MRA_Questions.SelectedIndex = tmpNewFocusRow
  else:
    # if -1 returned, this means that we are using 'groups' and we cannot get the 'index' without first refreshing DG
    # so here we'll need to iterate over DG items and select the items with the matching ID
    refresh_MRA_Questions(for_MRA_ID=tb_ThisMRAid.Text)
    dg_MRA_Questions.Focus()
    pCount = -1
    for xRow in dg_MRA_Questions.Items:
      pCount += 1
      if xRow['ID'] == selectedID:
        dg_MRA_Questions.SelectedIndex = pCount
        # and scroll into view
        dg_MRA_Questions.ScrollIntoView(xRow)
        break
  return
  
def MoveUp_MRA_Question(s, event):
  #! SEE Notes in Teams - need to add 'dgGroupColName' to this and similar functions to ensure we're only moving an item within the same grouping
  #! Also: need to update when adding/saving question details, if user selected a new 'group', we need to add to END of the GROUP
  # This function will move the selected Question up one row (and all other items down one)
  tmpNewFocusRow = dgItem_MoveUp(dgControl=dg_MRA_Questions, 
                                 tableToUpdate='Usr_MRA_TemplateQs', 
                                 sqlOrderColName='DisplayOrder', 
                                 dgIDcolName='RowID', 
                                 dgOrderColName='Order', 
                                 sqlOtherCheckCol='TypeID', 
                                 sqlOtherCheckValue=int(tb_ThisMRAid.Text), 
                                 sqlGroupColName='QGroupID',
                                 dgGroupColName='SectionID')
  if tmpNewFocusRow > -1: 
    refresh_MRA_Questions(for_MRA_ID=tb_ThisMRAid.Text)
    dg_MRA_Questions.Focus()
    dg_MRA_Questions.SelectedIndex = tmpNewFocusRow
  return
  
def MoveDown_MRA_Question(s, event):
  # This function will move the selected Question down one row (and all other items up one)
   #                dgItem_MoveDown(      dgControl,        tableToUpdate, sqlOrderColName, dgIDcolName, dgOrderColName, sqlOtherCheckCol, sqlOtherCheckValue)
  tmpNewFocusRow = dgItem_MoveDown(dgControl=dg_MRA_Questions, 
                                   tableToUpdate='Usr_MRA_TemplateQs', 
                                   sqlOrderColName='DisplayOrder', 
                                   dgIDcolName='RowID', 
                                   dgOrderColName='Order', 
                                   sqlOtherCheckCol='TypeID', 
                                   sqlOtherCheckValue=int(tb_ThisMRAid.Text), 
                                   sqlGroupColName='QGroupID',
                                   dgGroupColName='SectionID')  
  if tmpNewFocusRow > -1: 
    refresh_MRA_Questions(for_MRA_ID=tb_ThisMRAid.Text)
    dg_MRA_Questions.Focus()
    dg_MRA_Questions.SelectedIndex = tmpNewFocusRow  
  return
  
def MoveBottom_MRA_Question(s, event):
  # This function will move the selected Question to the bottom row (and all other items up one)
  selectedID = int(dg_MRA_Questions.SelectedItem['ID'])
  tmpNewFocusRow = dgItem_MoveToBottom(dgControl=dg_MRA_Questions, 
                                       tableToUpdate='Usr_MRA_TemplateQs', 
                                       sqlOrderColName='DisplayOrder', 
                                       dgIDcolName='RowID', 
                                       dgOrderColName='Order', 
                                       sqlOtherCheckCol='TypeID',
                                       sqlOtherCheckValue=int(tb_ThisMRAid.Text), 
                                       sqlGroupColName='QGroupID',
                                       dgGroupColName='SectionID')
  if tmpNewFocusRow > -1:
    refresh_MRA_Questions(for_MRA_ID=tb_ThisMRAid.Text)
    dg_MRA_Questions.Focus()
    dg_MRA_Questions.SelectedIndex = tmpNewFocusRow  
  else:
    # if -1 returned, this means that we are using 'groups' and we cannot get the 'index' without first refreshing DG
    # so here we'll need to iterate over DG items and select the items with the matching ID
    refresh_MRA_Questions(for_MRA_ID=tb_ThisMRAid.Text)
    dg_MRA_Questions.Focus()
    pCount = -1
    for xRow in dg_MRA_Questions.Items:
      pCount += 1
      if xRow['ID'] == selectedID:
        #dg_MRA_Questions.SelectedItem
        dg_MRA_Questions.SelectedIndex = pCount
        # and scroll into view
        dg_MRA_Questions.ScrollIntoView(xRow)
        break

  return
  
def Delete_MRA_Question(s, event):
  # This function will delete the selected Question (after confirmation)
  tmpID = lbl_QID.Content     #dg_MRA_Questions.SelectedItem['ID']
  tmpNewFocusRow = dgItem_DeleteSelected(dgControl=dg_MRA_Questions, 
                                         tableToUpdate='Usr_MRA_TemplateQs', 
                                         sqlOrderColName='DisplayOrder', 
                                         dgIDcolName='RowID', 
                                         dgOrderColName='Order', 
                                         dgNameDescColName='QText', 
                                         sqlOtherCheckCol='TypeID', 
                                         sqlOtherCheckValue=int(tb_ThisMRAid.Text), 
                                         sqlGroupColName='QGroupID',
                                         dgGroupColName='SectionID') 
  if tmpNewFocusRow > -1:
    refresh_MRA_Questions(for_MRA_ID=tb_ThisMRAid.Text)
    dg_MRA_Questions.Focus()
    dg_MRA_Questions.SelectedIndex = tmpNewFocusRow 
    
    # also need to delete Answers associated to Question
    deleteA_SQL = "DELETE FROM Usr_MRA_TemplateAs WHERE QuestionID = {0}".format(tmpID) 
    runSQL(deleteA_SQL, True, "There was an error deleting the Answers associated to the Questions used for the selected Matter Risk Assessment", "Error: Deleting Questions' Answers...")    
  return
  
def BackToOverview_MRA_Question(s, event):
  # This function should clear the 'Questions' tab and take us back to the 'Risk Assessment Overview' tab
  ti_MRA_Questions.Visibility = Visibility.Collapsed
  ti_MRA_Overview.Visibility = Visibility.Visible
  ti_MRA_Overview.IsSelected = True
  refresh_MRA_Templates(s, event)
  return
  
  
# # # #   END OF:   Q U E S T I O N S   # # # #

# # # #   P R E V I E W   M A T T E R   R I S K   A S S E S S M E N T   TAB # # # #

class preview_MRA(object):
  def __init__(self, myID, myOrder, myQuestion, myAnsGrp, myQGroup, myQID, 
               myAnswerID, myAnswerText, myScore, myEC, myQGroupName):
    self.RowID = myID
    self.DOrder = myOrder
    self.QuestionText = myQuestion
    self.AnswerGroupName = myAnsGrp
    self.QGroupID = myQGroup
    self.QuestionID = myQID
    self.LUP_AnswerID = myAnswerID
    self.LUP_AnswerText = '' if myAnswerID == -1 else myAnswerText
    self.LUP_Score = myScore
    self.LUP_EmailComment = myEC
    self.QGroupName = myQGroupName        #! New 04/09/2025
    return
    
  def __getitem__(self, index):
    # AnswerList    = Answer List GROUP name
    # Answer        = Answer ID (as per Usr_MRA_TemplateAs)
    # AText         = Looked up Answer Text (using AnswerID) AND tbAnswerText if AnswerID = -2 (TextBox)
    # SourceAnswers = List of Available Answers (broken up) - NEW: QuestionID (if greater than 0)
    if index == 'ID':
      return self.RowID
    elif index == 'Order':
      return self.DOrder
    elif index == 'Question':
      return self.QuestionText
    elif index == 'AnswerGroupName':
      return self.AnswerGroupName
    elif index == 'QGroupID':
      return self.QGroupID
    elif index == 'QGroupName':         #! New 04/09/2025
      return self.QGroupName
    elif index == 'Qid':
      return self.QuestionID
    elif index == 'AnswerID':
      return self.LUP_AnswerID
    elif index == 'AnswerText':
      return self.LUP_AnswerText
    elif index == 'Score':
      return self.LUP_Score
    elif index == 'EmailComment':
      return self.LUP_EmailComment
    else:
      return ''
      
def refresh_Preview_MRA(s, event):
  # This function will populate the Matter Risk Assessment Preview datagrid
  #MessageBox.Show("Start - getting group ID", "Refreshing list (datagrid of questions)")
  #dg_MRAPreview.ItemsSource = None
  
  tmpGroup = ''
  if dg_GroupItems_Preview.SelectedIndex > -1:
    if dg_GroupItems_Preview.SelectedItem['ID'] != 0:
      tmpGroup = dg_GroupItems_Preview.SelectedItem['ID']
      showGrouping = False
    else:
      showGrouping = True

  #MessageBox.Show("Genating SQL...", "Refreshing list (datagrid of questions)")
  mySQL = """SELECT '0-RowID' = MRAP.ID, '1-DOrder' = MRAP.DOrder, '2-QuestionText' = MRAP.QuestionText, '3-AnswerGroupName' = MRAP.AnswerList, 
                '4-QGroupID' = MRAP.QGroupID, '5-QuestionID' = MRAP.QuestionID, '6-LUP_AnswerID' = MRAP.AnswerID, 
                '7-LUP_AnswerText' = CASE WHEN MRAP.AnswerList = '(TextBox)' THEN tbAnswerText ELSE (SELECT AnswerText FROM Usr_MRA_TemplateAs WHERE AnswerID = MRAP.AnswerID AND QuestionID = MRAP.QuestionID) END, 
                '8-LUP_Score' = MRAP.Score, '9-LUP_EmailComment' = ISNULL(EmailComment, ''),
                '10-QGroupN' = QG.Name 
            FROM Usr_MRA_Preview MRAP 
              LEFT OUTER JOIN Usr_MRA_QGroups QG ON MRAP.QGroupID = QG.ID """
  
  if tmpGroup != '':
    mySQL += "WHERE MRAP.QGroupID = {0} ".format(tmpGroup)
  
  # add order
  mySQL += "ORDER BY QG.DisplayOrder, MRAP.DOrder"
  #MessageBox.Show("SQL: " + str(mySQL) + "\n\nRefreshing list (datagrid of questions)", "Debug: Populating List of Questions (Preview MRA)")

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
          iAnsGrpName = '' if dr.IsDBNull(3) else dr.GetString(3)
          iQGrpID = 0 if dr.IsDBNull(4) else dr.GetValue(4)
          iQid = 0 if dr.IsDBNull(5) else dr.GetValue(5)
          iAnsID = '' if dr.IsDBNull(6) else dr.GetValue(6)
          iLUP_AnsTxt = '' if dr.IsDBNull(7) else dr.GetString(7)
          iLUP_AnsScore = '' if dr.IsDBNull(8) else dr.GetValue(8)
          iAEC = '' if dr.IsDBNull(9) else dr.GetString(9)
          iQGrpName = '' if dr.IsDBNull(10) else dr.GetString(10)

          myItems.append(preview_MRA(myID=iID, myOrder=iDO, myQuestion=iQText, myAnsGrp=iAnsGrpName, 
                                     myQGroup=iQGrpID, myQID=iQid, myAnswerID=iAnsID, myAnswerText=iLUP_AnsTxt,
                                     myScore=iLUP_AnsScore, myEC=iAEC, myQGroupName=iQGrpName))  
      
    dr.Close()
  _tikitDbAccess.Close()
  
  #MessageBox.Show("Putting list items into Datagrid", "Refreshing list (datagrid of questions)")
  #! New: 04/09/2025 - added 'Grouing' to the DataGrid, now we've sorted the 'DisplayOrder' - this shold only show when 'View All' option selected
  if showGrouping == False: 
    dg_MRAPreview.ItemsSource = myItems
  else:
    tmpC = ListCollectionView(myItems)
    tmpC.GroupDescriptions.Add(PropertyGroupDescription('QGroupName'))
    dg_MRAPreview.ItemsSource = tmpC
  return


def MRA_Preview_AutoAdvance(currentDGindex, s, event):
  
  #MessageBox.Show("Test - do we get here?")    # debug
  currPlusOne = currentDGindex + 1
  totalDGitems = dg_MRAPreview.Items.Count
  #MessageBox.Show("Current Index: " + str(currentDGindex) + "\nPlusOne: " + str(currPlusOne) + "\nTotalDGitems: " + str(totalDGitems), "Auto-Advance to next question...")
  
  # if we're at the end of current list...
  if currPlusOne == totalDGitems:
    # get group details...
    currGroupIndex = dg_GroupItems_Preview.SelectedIndex
    totalGroupItems = dg_GroupItems_Preview.Items.Count  
    
    # if at end of current question group...
    if currGroupIndex == (totalGroupItems - 1):
      # select '0' (show all)
      dg_GroupItems_Preview.SelectedIndex = 0
    else:
      # then select next 'group'
      dg_GroupItems_Preview.SelectedIndex = currGroupIndex + 1
    dg_MRAPreview.SelectedIndex = 0
    
  else:
    dg_MRAPreview.SelectedIndex = currPlusOne
    dg_MRAPreview.ScrollIntoView(dg_MRAPreview.Items[currPlusOne])
  return
  

def MRA_Preview_UpdateTotalScore(s, event):
  # This function will update the overall total score
  tmpTotal = int(_tikitResolver.Resolve("[SQL: SELECT SUM(Score) FROM Usr_MRA_Preview]"))
  lbl_MRAPreview_Score.Content = str(tmpTotal)     #'{:,.0f}'.format(tmpTotal)
  
  tmpSQL_ID = "[SQL: SELECT SM.LMH_ID FROM Usr_MRA_ScoreMatrix SM WHERE SM.TypeID = {0} AND ({1} BETWEEN SM.Score_From AND SM.Score_To)]".format(lbl_MRA_Preview_ID.Content, tmpTotal)
  try:
    tmpCatID = _tikitResolver.Resolve(tmpSQL_ID)
    lbl_MRAPreview_RiskCategoryID.Content = str(tmpCatID)
  except:
    MessageBox.Show("There was an error getting the Low, Medium or High ID, using SQL:\n" + tmpSQL_ID, "Error: Getting Low, Med, High ID")

  tmpSQL_Text = """[SQL: SELECT 'LMH Text' = CASE SM.LMH_ID WHEN  1 THEN 'Low' WHEN 2 THEN 'Medium' WHEN 3 THEN 'High' END 
                         FROM Usr_MRA_ScoreMatrix SM 
                         WHERE SM.TypeID = {0} AND ({1} BETWEEN SM.Score_From AND SM.Score_To)]""".format(lbl_MRA_Preview_ID.Content, tmpTotal)
  try:
    tmpCat = _tikitResolver.Resolve(tmpSQL_Text)
    lbl_MRAPreview_RiskCategory.Content = tmpCat
  except:
    MessageBox.Show("There was an error getting the Low, Medium or High text, using SQL:\n" + tmpSQL_Text, "Error: Getting Low, Med, High Text")
    
  return


def MRA_Preview_SelectionChanged(s, event):
  if dg_MRAPreview.SelectedIndex > -1:
    lbl_MRAPreview_DGID.Content = dg_MRAPreview.SelectedItem['ID']
    lbl_MRAPreview_CurrVal.Content = dg_MRAPreview.SelectedItem['AnswerText']
    tb_previewMRA_QestionText.Text = dg_MRAPreview.SelectedItem['Question']
    tb_MRAPreview_EC.Text = dg_MRAPreview.SelectedItem['EmailComment']
    
    if dg_MRAPreview.SelectedItem['AnswerGroupName'] == '(TextBox)':
      # here we need to display the text box and hide the combo box...
      #cbo_preview_MRA_SelectedComboAnswer.SelectedIndex = -1
      cbo_preview_MRA_SelectedComboAnswer.Visibility = Visibility.Collapsed
      tb_preview_MRA_SelectedTextAnswer.Text = dg_MRAPreview.SelectedItem['AnswerText']
      tb_preview_MRA_SelectedTextAnswer.Visibility = Visibility.Visible
      update_EmailComment(s, event)
    else:
      # Anything else is considered a drop-down (combo box), so hide the text box and display the combo box...
      tb_preview_MRA_SelectedTextAnswer.Text = ''
      tb_preview_MRA_SelectedTextAnswer.Visibility = Visibility.Collapsed
      
      cbo_preview_MRA_SelectedComboAnswer.Visibility = Visibility.Visible
      populate_MRA_Preview_SelectAnswerCombo(s, event)
      # select 'nothing'...
      cbo_preview_MRA_SelectedComboAnswer.SelectedIndex = -1
      
      # now select the appropriate drop-down item if answered previously
      pCount = -1
      for xRow in cbo_preview_MRA_SelectedComboAnswer.Items:
        pCount += 1
        #MessageBox.Show("xRow: " + str(xRow) + "\nAText: " + str(dg_MRAPreview.SelectedItem['AnswerText']))
        if xRow.AText == dg_MRAPreview.SelectedItem['AnswerText']:
          cbo_preview_MRA_SelectedComboAnswer.SelectedIndex = pCount
          break
      
    btn_preview_MRA_SaveAnswer.IsEnabled = True
    
  else:
    lbl_MRAPreview_DGID.Content = ''
    lbl_MRAPreview_CurrVal.Content = ''    
    tb_previewMRA_QestionText.Text = '-NO QUESTION SELECTED - PLEASE SELECT FROM THE LIST ABOVE-'
    btn_preview_MRA_SaveAnswer.IsEnabled = False
    tb_preview_MRA_SelectedTextAnswer.Text = ''
    tb_preview_MRA_SelectedTextAnswer.Visibility = Visibility.Collapsed
    cbo_preview_MRA_SelectedComboAnswer.Visibility = Visibility.Collapsed
    tb_MRAPreview_EC.Text = ''
    #cbo_preview_MRA_SelectedComboAnswer.SelectedIndex = -1
    #cbo_preview_MRA_SelectedComboAnswer.ItemsSource = None
    # Hmm... point to note: keep getting errors when setting 'SelectedIndex' to -1
    
  return
  

class YourAnswer(object):
  def __init__(self, myText, myScore, myEC, myAid):
    self.AText = myText
    self.AScore = myScore
    self.AEmailComment = myEC
    self.AID = myAid
    return
    
  def __getitem__(self, index):
    if index == 'Text':
      return self.AText
    elif index == 'Score':
      return self.AScore
    elif index == 'EmailComment':
      return self.AEmailComment
    elif index == 'ID':
      return self.AID
    else:
      return ''

def populate_MRA_Preview_SelectAnswerCombo(s, event):
  # New 2nd May 2024 - this will populate the Combo box on the 'Preview MRA' tab for the selected Question
  
  if dg_MRAPreview.SelectedIndex > -1:
    tmpQID = dg_MRAPreview.SelectedItem['Qid']
  else:
    return
  
  mySQL = "SELECT AnswerText, Score, EmailComment, AnswerID FROM Usr_MRA_TemplateAs WHERE QuestionID = {0} ORDER BY DisplayOrder".format(tmpQID)

  _tikitDbAccess.Open(mySQL)
  myItems = []
  
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          iAText = '' if dr.IsDBNull(0) else dr.GetString(0)
          iScore = 0 if dr.IsDBNull(1) else dr.GetValue(1)
          iEmailComment = '' if dr.IsDBNull(2) else dr.GetString(2)
          iID = 0 if dr.IsDBNull(3) else dr.GetValue(3)
          
          myItems.append(YourAnswer(iAText, iScore, iEmailComment, iID))
      
    dr.Close()
  _tikitDbAccess.Close()

  cbo_preview_MRA_SelectedComboAnswer.ItemsSource = myItems
  return


def PreviewMRA_BackToOverview(s, event):
  ti_MRA_Overview.Visibility = Visibility.Visible
  ti_MRA_Preview.Visibility = Visibility.Collapsed
  ti_MRA_Overview.IsSelected = True
  return



# New Question Groups Section
class QGroups(object):
  def __init__(self, myID, myDO, myDesc, myTotalQs):
    self.gID = myID
    self.gDO = myDO
    self.gDesc = myDesc
    self.gTotalQs = myTotalQs
    return
    
  def __getitem__(self, index):
    if index == 'ID':
      return self.gID
    elif index == 'Order':
      return self.gDO
    elif index == 'Desc':
      return self.gDesc     
    elif index == 'TotalQs':
      return self.gTotalQs 
    else:
      return ''
      
def populate_preview_MRA_QGroups(s, event):
  # This function populates the new Question Groups data grid on the Preview MRA tab
  mySQL = """SELECT QG.ID, QG.DisplayOrder, QG.Name, COUNT(MRAP.ID) FROM Usr_MRA_Preview MRAP 
            LEFT OUTER JOIN Usr_MRA_QGroups QG ON MRAP.QGroupID = QG.ID 
            GROUP BY QG.ID, QG.DisplayOrder, QG.Name ORDER BY QG.DisplayOrder"""
  
  try:
    totalQCount = _tikitResolver.Resolve("[SQL: SELECT COUNT(ID) FROM Usr_MRA_Preview]")
  except:
    totalQCount = 0

  myItems = []
  myItems.append(QGroups(0, 0, '(View All)', totalQCount))

  if int(totalQCount) > 0: 
    _tikitDbAccess.Open(mySQL)
  
    if _tikitDbAccess._dr is not None:
      dr = _tikitDbAccess._dr
      if dr.HasRows:
        while dr.Read():
          if not dr.IsDBNull(0):
            iID = 0 if dr.IsDBNull(0) else dr.GetValue(0)
            iDO = 0 if dr.IsDBNull(1) else dr.GetValue(1)
            iQGroup = '-' if dr.IsDBNull(2) else dr.GetString(2)
            iCount = 0 if dr.IsDBNull(3) else dr.GetValue(3)
                    
            myItems.append(QGroups(iID, iDO, iQGroup, iCount))  
      
      dr.Close()
    _tikitDbAccess.Close()
  
  dg_GroupItems_Preview.ItemsSource = myItems
  
  if dg_GroupItems_Preview.Items.Count == 1:
    grid_Preview_MRA.Visibility = Visibility.Hidden
    tb_NoMRA_PreviewQs.Visibility = Visibility.Visible
  else:
    grid_Preview_MRA.Visibility = Visibility.Visible
    tb_NoMRA_PreviewQs.Visibility = Visibility.Hidden
    #dg_GroupItems_Preview.SelectedIndex = 0
  return  
  
  
def preview_MRA_SaveAnswer(s, event):
  # This replaces the 'Cell Edit Ending' function of the (old) editable data grid (now putting values beneath DG)

  # get current values
  rowID = dg_MRAPreview.SelectedItem['ID']
  tmpEC = tb_MRAPreview_EC.Text   # Commented out other lines that look this up, as now doing that via updating combo box (or text box), so just need to get current value here
  tmpQID = dg_MRAPreview.SelectedItem['Qid']
  updateSQL = "[SQL: UPDATE Usr_MRA_Preview SET "
  
  if dg_MRAPreview.SelectedItem['AnswerGroupName'] == '(TextBox)':
    # Answers with a value of -2 denote a TEXT BOX answer...
    newTextVal = tb_preview_MRA_SelectedTextAnswer.Text
    if len(newTextVal.strip()) == 0:
      # text string is empty, set answer to 'no answer' score
      newTextVal = ""
      tmpScore = _tikitResolver.Resolve("[SQL: SELECT Score FROM Usr_MRA_TemplateAs WHERE GroupName = '(TextBox)' AND AnswerText = '(TextBox-Empty)' AND QuestionID = {0}]".format(tmpQID))
      tmpAnsID = _tikitResolver.Resolve("[SQL: SELECT ID FROM Usr_MRA_TemplateAs WHERE GroupName = '(TextBox)' AND AnswerText = '(TextBox-Empty)' AND QuestionID = {0}]".format(tmpQID))
    else:
      # text string NOT empty, get score for not empty
      newTextVal = newTextVal.replace("'", "''")
      tmpScore = _tikitResolver.Resolve("[SQL: SELECT Score FROM Usr_MRA_TemplateAs WHERE GroupName = '(TextBox)' AND AnswerText = '(TextBox-NotEmpty)' AND QuestionID = {0}]".format(tmpQID))
      tmpAnsID = _tikitResolver.Resolve("[SQL: SELECT ID FROM Usr_MRA_TemplateAs WHERE GroupName = '(TextBox)' AND AnswerText = '(TextBox-NotEmpty)' AND QuestionID = {0}]".format(tmpQID))

    #tmpAnsID = -2
    updateSQL += "tbAnswerText = '{0}', ".format(newTextVal)
    
  else:  
    # All other answers relate to the combo box...
    # if nothing selected in combo box...
    if cbo_preview_MRA_SelectedComboAnswer.SelectedIndex == -1:
      tmpAnsID = -1
      tmpScore = 0
    else:
      # lookup answer index, score, and EmailComment
      newTextVal = cbo_preview_MRA_SelectedComboAnswer.SelectedItem['Text']
      tmpAnsID = cbo_preview_MRA_SelectedComboAnswer.SelectedItem['ID']
      tmpScore = cbo_preview_MRA_SelectedComboAnswer.SelectedItem['Score']
      #tmpEC = cbo_preview_MRA_SelectedComboAnswer.SelectedItem['EmailComment']

  #MessageBox.Show("rowID: " + str(rowID) + "\nNewTextVal: " + str(newTextVal) + "\nFromAnsList: " + str(fromAnsList))
  #MessageBox.Show("tmpAnsID: " + str(tmpAnsID) + "\ntmpScore: " + str(tmpScore))
  if len(tmpEC) > 0:
    tmpEC = tmpEC.replace("'", "''")

  updateSQL += "AnswerID = {0}, Score = {1}, EmailComment = '{2}' WHERE ID = {3}]".format(tmpAnsID, tmpScore, tmpEC, rowID)
  canContinue = False
  try:
    _tikitResolver.Resolve(updateSQL)
    canContinue = True
  except:
    MessageBox.Show("There was an error updating the answer (no updates made!), using SQL:\n" + updateSQL, "Error: MRA Preview - Updating Answer...")
    
  if canContinue == True:
    # need to get current index as 'refresh' will wipe out all items...
    currDGindex = dg_MRAPreview.SelectedIndex
    MRA_Preview_UpdateTotalScore(s, event) 
    refresh_Preview_MRA(s, event)
    
    if chk_MRAPreview_AutoSelectNext.IsChecked == True:
      MRA_Preview_AutoAdvance(currDGindex, s, event)
  return
  

def GroupItems_Preview_SelectionChanged(s, event):
  refresh_Preview_MRA(s, event)
  dg_MRAPreview.SelectedIndex = 0
  #MessageBox.Show("This is a test to see if selection change happened", "DEBUG: Test GroupItems_Preview_SelectionChanged")
  return
  

def update_EmailComment(s, event):
  if dg_MRAPreview.SelectedIndex == -1:
    return
    
  rowID = dg_MRAPreview.SelectedItem['ID']
  tmpQID = dg_MRAPreview.SelectedItem['Qid']
  
  if dg_MRAPreview.SelectedItem['AnswerGroupName'] == '(TextBox)':
    # Answers with a value of -2 denote a TEXT BOX answer...
    newTextVal = tb_preview_MRA_SelectedTextAnswer.Text
    newTextVal = newTextVal.replace("'", "''")
    
    if len(newTextVal.strip()) == 0:
      # text string is empty, set answer to 'no answer' score
      tmpEC = _tikitResolver.Resolve("[SQL: SELECT ISNULL(EmailComment, '') FROM Usr_MRA_TemplateAs WHERE GroupName = '(TextBox)' AND AnswerText = '(TextBox-Empty)' AND QuestionID = {0}]".format(tmpQID))
    else:
      tmpEC = _tikitResolver.Resolve("[SQL: SELECT ISNULL(EmailComment, '') FROM Usr_MRA_TemplateAs WHERE GroupName = '(TextBox)' AND AnswerText = '(TextBox-NotEmpty)' AND QuestionID = {0}]".format(tmpQID))
  
  else:
    # All other answers relate to the combo box...
    if cbo_preview_MRA_SelectedComboAnswer.SelectedIndex == -1:
      tmpEC = ''
    else:
      tmpEC = cbo_preview_MRA_SelectedComboAnswer.SelectedItem['EmailComment']
      #newTextVal = cbo_preview_MRA_SelectedComboAnswer.SelectedItem['Text']
      #fromAnsList = dg_MRAPreview.SelectedItem['AnswerList']
      #tmpEC = _tikitResolver.Resolve("[SQL: SELECT ISNULL(EmailComment, '') FROM Usr_MRA_TemplateAs WHERE GroupName = '" + str(fromAnsList) + "' AND AnswerText = '" + str(newTextVal) + "' AND QuestionID = " + str(tmpQID) + "]")

  tb_MRAPreview_EC.Text = tmpEC
  return
  
# # # #   END OF:   P R E V I E W   M A T T E R   R I S K   A S S E S S M E N T   TAB # # # #



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
  cbo_QuestionAnswerList.ItemsSource = tmpItem2
  
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
  populate_MRA_QGroups()
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

def updateAnswerList_forSelectedQuestion(s, event):
  #! [Configure MRA] tab > [Editing Questions] tab > 'Update' button clicked in the 'Answer Options' area
  # Here we need to first check if any items currently exist in list... if so, ask whether we wish to clear original list first (if no, we'll append to end if applicable)
  noOfInserts = 1
  tmpQID = str(lbl_QID.Content)
  if tmpQID == "":
    return

  # check inputs are valid
  if opt_CopyAnswersFrom.IsChecked == True:
    if cbo_QuestionAnswerList.SelectedIndex == -1:
      MessageBox.Show("You haven't selected a source list to copy items from!\n\nPlease select an item from the adjacent drop-down before continuing", "Update Answer List for Question...")
      return

  if opt_CopyAnswersFrom.IsChecked == False and opt_BlankAnswerList.IsChecked == False:
    MessageBox.Show("You haven't selected whether to copy items from existing list, or to start with a blank list!\n\nPlease select an Answer Option before continuing", "Update Answer List for Question...")
    return

  #! New 25/07/2025 - getting the NEXT available 'AnswerID' to use for new items added (we add on initial insert now rather than doing at end of function)
  #! NB: we don't '+1' here as we're going to just add the 'DisplayOrder' to this 'max' number - 'DisplayOrder' starts at 1
  nextAnsID = _tikitResolver.Resolve("[SQL: SELECT ISNULL(MAX(AnswerID), 1) FROM Usr_MRA_TemplateAs]")

  # get count of items currently in list (helpful for later on in order to set DisplayOrder correctly)
  noOfItemsInList = int(_tikitResolver.Resolve("[SQL: SELECT COUNT(ID) FROM Usr_MRA_TemplateAs WHERE QuestionID = {0}]".format(tmpQID)))

  # now check if any items currently exist...
  if int(noOfItemsInList) > 0:
    tmpMsg = "Would you like to overwrite current list of answers?\n\nYes - will delete current answers and import from other list\nNo - existing answers remain, other list items added to end"
    myResult = MessageBox.Show(tmpMsg, "Overwrite current items...", MessageBoxButtons.YesNo)
    
    if myResult == DialogResult.Yes:
      # delete existing Answers for this Question
      deleteSQL = "DELETE FROM Usr_MRA_TemplateAs WHERE QuestionID = " + str(tmpQID)
      runSQL(deleteSQL, True, "There was an error clearing the original items from the list", "Update Answer Options - Delete Original List...")
      # re-get count of items in list (doing this instead of assuming above delete worked correctly)
      noOfItemsInList = int(_tikitResolver.Resolve("[SQL: SELECT COUNT(ID) FROM Usr_MRA_TemplateAs WHERE QuestionID = {0}]".format(tmpQID)))      

  tmpUseAltResolver = False  
  # check inputs are valid
  if opt_CopyAnswersFrom.IsChecked == True:
    # should set the SQL here for the selected 'Answer List' (in combo box)
    tmpUseAltResolver = True
    insSQL = """WITH myCopyList AS (
                    SELECT 'GrpName' = GroupName, 'QID' = {currQID}, 'AnsText' = AnswerText, 'Score' = Score, 
                        'EmailComment' = EmailComment, 'RowNum' = ROW_NUMBER() OVER (ORDER BY DisplayOrder, AnswerText),  
                        'DispOrder' = DisplayOrder 
                    FROM Usr_MRA_TemplateAs WHERE QuestionID = {QtoCopyID} 
                  )
                INSERT INTO Usr_MRA_TemplateAs (GroupName, QuestionID, AnswerText, Score, EmailComment, DisplayOrder, AnswerID) 
                SELECT GrpName, QID, AnsText, Score, EmailComment, RowNum, {nextAnswerID} + RowNum FROM myCopyList ORDER BY RowNum;
                """.format(currQID=tmpQID, QtoCopyID=cbo_QuestionAnswerList.SelectedItem['ID'], nextAnswerID=nextAnsID)
    tmpNameAnsList1 = cbo_QuestionAnswerList.SelectedItem['Name']
    tmpNameAnsList = tmpNameAnsList1.replace("(GLOBAL) ", "")
    updateQ_SQL = "UPDATE Usr_MRA_TemplateQs SET AnswerList = '{0}' WHERE QuestionID = {1}".format(tmpNameAnsList, tmpQID)
    
  elif opt_BlankAnswerList.IsChecked == True:
    noOfInserts = 2
    insSQL = "INSERT INTO Usr_MRA_TemplateAs (QuestionID, AnswerText, Score, EmailComment, DisplayOrder, AnswerID) VALUES ({0}, 'Answer1', 0, '', {1} + 1, {2} + 1)".format(tmpQID, noOfItemsInList, nextAnsID)
    insSQL2 = "INSERT INTO Usr_MRA_TemplateAs (QuestionID, AnswerText, Score, EmailComment, DisplayOrder, AnswerID) VALUES ({0}, 'Answer2', 0, '', {1} + 2, {2} + 2)".format(tmpQID, noOfItemsInList, nextAnsID)
    updateQ_SQL = "UPDATE Usr_MRA_TemplateQs SET AnswerList = '' WHERE QuestionID = {0}".format(tmpQID)
  
  # now add applicable items...
  runSQL(codeToRun=insSQL, showError=True, 
         errorMsgText="There was an error adding the Answer list...", errorMsgTitle="Error: adding Answer list...", 
         useAltResolver=tmpUseAltResolver)
  if noOfInserts == 2:
    runSQL(codeToRun=insSQL2, showError=True, errorMsgText="There was an error adding the Answer list...", errorMsgTitle="Error: adding Answer list...")
  
  # run code to rename the list (to remove the 'Global' as we've now copied down to this Question)
  runSQL(codeToRun=updateQ_SQL, showError=True, errorMsgText="There was an error updating the Questions 'AnswerList' field", errorMsgTitle="Error: Updating 'AnswerList' for Question...")
  
  # finally update TemplateAnswers AnswerID to match ID where item is null
  #runSQL("UPDATE Usr_MRA_TemplateAs SET AnswerID = ID WHERE AnswerID IS NULL", False, '', '')
  #! 1) Why is this not just done on initial insert? 2) Why TF are we copying ID to the 'AnswerID' field??
  #! This really out to just get the MAX number of 'AnswerID' and +1 to each new answer item added
  
  # finally refresh the answer list...
  populate_AnswersPreview(s, event)
  return

class AnswersListPreview(object):
  def __init__(self, myDO, myQA_Text, myQA_Scrore, myEC, myID, myRowID):
    self.mraQA_DO = myDO
    self.mraQA_Text = myQA_Text
    self.mraQA_Score = myQA_Scrore
    self.mraQA_EC = myEC
    self.mraQA_ID = myID
    self.mraQA_RowID = myRowID
    return
    
  def __getitem__(self, index):
    if index == 'Order':
      return self.mraQA_DO
    elif index == 'AText':
      return self.mraQA_Text
    elif index == 'AScore':
      return self.mraQA_Score
    elif index == 'EmailComment':
      return self.mraQA_EC
    elif index == 'ID':
      return self.mraQA_ID
    elif index == 'RowID':
      return self.mraQA_RowID
    else:
      return ''

def populate_AnswersPreview(s, event):
  #! [Configure MRA] tab > [Editing Questions] tab > DataGrid in the 'Answer Options' area
  # Populates the list of Answers for the selected Question on the 'Editing Questions' tab of a MRA
  
  # form SQL to get all answers for the currently selected question
  getTableSQL = """SELECT MRATA.AnswerID, MRATA.DisplayOrder, MRATA.AnswerText, MRATA.Score, MRATA.EmailComment, MRATA.ID  
                  FROM Usr_MRA_TemplateAs MRATA 
                  WHERE MRATA.QuestionID = {0}
                  ORDER BY MRATA.DisplayOrder""".format(lbl_QID.Content)
  
  tmpItem = []
  _tikitDbAccess.Open(getTableSQL)
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          tmpID = 0 if dr.IsDBNull(0) else dr.GetValue(0)
          tmpDO = 0 if dr.IsDBNull(1) else dr.GetValue(1)
          tmpAText = '' if dr.IsDBNull(2) else dr.GetString(2)
          tmpScore = 0 if dr.IsDBNull(3) else dr.GetValue(3)
          tmpEC = '' if dr.IsDBNull(4) else dr.GetString(4)
          tmpRowID = 0 if dr.IsDBNull(5) else dr.GetValue(5)

          tmpItem.append(AnswersListPreview(myDO=tmpDO, myQA_Text=tmpAText, myQA_Scrore=tmpScore, myEC=tmpEC, myID=tmpID, myRowID=tmpRowID))

    dr.Close()
  
  #close db connection
  _tikitDbAccess.Close()
  dg_EditMRA_AnswersPreview.ItemsSource = tmpItem
  
  if dg_EditMRA_AnswersPreview.Items.Count == 0:
    dg_EditMRA_AnswersPreview.Visibility = Visibility.Collapsed
    lbl_NoAnswers.Visibility = Visibility.Visible
    btn_CopySelectedListItem1.Visibility = Visibility.Collapsed
    btn_A_MoveTop1.Visibility = Visibility.Collapsed
    btn_A_MoveUp1.Visibility = Visibility.Collapsed
    btn_A_MoveDown1.Visibility = Visibility.Collapsed
    btn_A_MoveBottom1.Visibility = Visibility.Collapsed
    btn_DeleteSelectedListItem1.Visibility = Visibility.Collapsed
    MRA_A_Sep1.Visibility = Visibility.Collapsed
    MRA_A_Sep2.Visibility = Visibility.Collapsed
    MRA_A_Sep3.Visibility = Visibility.Collapsed
    MRA_A_Sep4.Visibility = Visibility.Collapsed
    MRA_A_Sep5.Visibility = Visibility.Collapsed
    MRA_A_Sep6.Visibility = Visibility.Collapsed
  else:
    dg_EditMRA_AnswersPreview.Visibility = Visibility.Visible
    lbl_NoAnswers.Visibility = Visibility.Collapsed
    btn_CopySelectedListItem1.Visibility = Visibility.Visible
    btn_A_MoveTop1.Visibility = Visibility.Visible
    btn_A_MoveUp1.Visibility = Visibility.Visible
    btn_A_MoveDown1.Visibility = Visibility.Visible
    btn_A_MoveBottom1.Visibility = Visibility.Visible
    btn_DeleteSelectedListItem1.Visibility = Visibility.Visible
    MRA_A_Sep1.Visibility = Visibility.Visible
    MRA_A_Sep2.Visibility = Visibility.Visible
    MRA_A_Sep3.Visibility = Visibility.Visible
    MRA_A_Sep4.Visibility = Visibility.Visible
    MRA_A_Sep5.Visibility = Visibility.Visible
    MRA_A_Sep6.Visibility = Visibility.Visible
  return


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


def populate_MRA_QGroups():
  # This function will populate the 'Section / Group' combo box on the 'Edit MRA Questions' tab

  mySQL = "SELECT DisplayOrder, ID, Name FROM Usr_MRA_QGroups WHERE ForWhat = 'MRA' ORDER BY DisplayOrder "
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

          tmpDDL.append(twoColList(tmpID, tmpDesc)) 
    dr.Close()
  
  #close db connection
  _tikitDbAccess.Close()

  cbo_QuestionGroup.ItemsSource = tmpDDL
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
ti_MRA_Questions = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_MRA_Questions')
ti_MRA_Overview = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_MRA_Overview')
ti_MRA_Preview = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_MRA_Preview')
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
cbo_GroupBy.SelectionChanged += refresh_ListOfLockedMatters
btn_RefreshLockedMatters = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_RefreshLockedMatters')
btn_RefreshLockedMatters.Click += refresh_ListOfLockedMatters
btn_ClearLockedMattersFilters = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_ClearLockedMattersFilters')
btn_ClearLockedMattersFilters.Click += clear_LockedMatters_Filters
btn_UnlockMatters = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_UnlockMatters')
btn_UnlockMatters.Click += unlockMatter
#btn_LockMatters = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_LockMatters')
#btn_LockMatters.Click += 
dg_LockedMatters = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_LockedMatters')
tb_NoLockedMatters = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_NoLockedMatters')
cbo_Reason = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cbo_Reason')


## C O N F I G U R E   M A T T E R   R I S K   A S S E S S M E N T    - TAB ##
btn_AddNew_MRATemplate = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_AddNew_MRATemplate')
btn_AddNew_MRATemplate.Click += AddNew_MRA_Template
btn_CopySelected_MRATemplate = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_CopySelected_MRATemplate')
btn_CopySelected_MRATemplate.Click += btn_Duplicate_MRA_Template
btn_DeleteSelected_MRATemplate = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_DeleteSelected_MRATemplate')
btn_DeleteSelected_MRATemplate.Click += Delete_MRA_Template
btn_Preview_MRATemplate = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Preview_MRATemplate')
btn_Preview_MRATemplate.Click += Preview_MRA_Template
btn_Edit_MRATemplate = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Edit_MRATemplate')
btn_Edit_MRATemplate.Click += btn_Edit_MRATemplate_Click
dg_MRA_Templates = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_MRA_Templates')
dg_MRA_Templates.SelectionChanged += DG_MRA_Template_SelectionChanged
#dg_MRA_Templates.CellEditEnding += DG_MRA_Template_CellEditEnding
chk_ShowHiddenNMRAtemplates = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'chk_ShowHiddenNMRAtemplates')
chk_ShowHiddenNMRAtemplates.Click += refresh_MRA_Templates
#^ removing the above
chk_ShowAllVersions = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'chk_ShowAllVersions')
chk_ShowAllVersions.Click += refresh_MRA_Templates

#! New fields added 21/08/2025
tb_Sel_MRA_Name = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_Sel_MRA_Name')            # TextBox for editing the name of the selected NMRA template 
lbl_Sel_MRA_ID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_Sel_MRA_ID')              # Label to store ID of selected NMRA template
#lbl_Sel_MRA_Name = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_Sel_MRA_Name')          # Label to display the ORIGINAL name of selected NMRA template
lbl_Sel_MRA_EditingTypeID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_Sel_MRA_EditingTypeID')  # Label to store the ID of the template this is the 'edit' version (if applicable)
tb_ExpiresInXdays = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_ExpiresInXdays')

tb_InternalNote = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_InternalNote')
tb_UsersNote = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_UsersNote')
chk_MRA_Hidden = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'chk_MRA_Hidden')
dtp_MRA_EffectiveFrom = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dtp_MRA_EffectiveFrom')
dtp_MRA_EffectiveTo = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dtp_MRA_EffectiveTo')
btn_Save_Main_MRA_Header_Details = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Save_Main_MRA_Header_Details')
btn_Save_Main_MRA_Header_Details.Click += btn_SaveMainMRA_Details_Click


## MRA SCORE THRESHOLDS ##
#lbl_Sel_MRA_Name = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_Sel_MRA_Name')
lbl_Sel_MRA_ID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_Sel_MRA_ID')
tb_ST_NoQs = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_ST_NoQs')
tb_ST_NoMRA_Selected = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_ST_NoMRA_Selected')
tb_SM_Low_To = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_SM_Low_To')
tb_SM_Med_To = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_SM_Med_To')
sld_Low_To = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'sld_Low_To')
sld_Low_To.ValueChanged += ST_Low_SliderChanged
sld_Med_To = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'sld_Med_To')
sld_Med_To.ValueChanged += ST_Med_SliderChanged
lbl_SM_High_To = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_SM_High_To')
lbl_SM_Med_From = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_SM_Med_From')
lbl_SM_High_From = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_SM_High_From')
grd_ScoreThresholds = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'grd_ScoreThresholds')
btn_SaveScoreThresholds = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_SaveScoreThresholds')
btn_SaveScoreThresholds.Click += btn_SaveScoreThresholds_Click 
btn_GetMax = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_GetMax')
btn_GetMax.Click += scoreMatrix_setMax
stk_ST_SelectedMRA = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'stk_ST_SelectedMRA')

## CASE TYPE  MAPPING  TO  MATTER RISK ASSESSMENT  TEMPLATE ##
dg_MRA_Templates_CTD = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_MRA_Templates_CTD')
#dg_MRA_Templates_CTD.SelectionChanged += MRA_Department_Defaults_SelectionChanged
dg_MRA_CaseTypes_MRATemplate = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_MRA_CaseTypes_MRATemplate')
dg_MRA_CaseTypes_MRATemplate.SelectionChanged += MRA_CaseType_Defaults_SelectionChanged

lbl_SelectedCaseType = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_SelectedCaseType')
#tog_ExpandContract = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tog_ExpandContract')
#tog_ExpandContract.Click += tog_ExpandOrContract
#exp = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'exp')

lbl_SelectedCaseTypeID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_SelectedCaseTypeID')
lbl_SelectedDeptName = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_SelectedDeptName')
lbl_SelectedDeptID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_SelectedDeptID')
btn_Save_MRA_TemplateToUseForCaseType = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Save_MRA_TemplateToUseForCaseType')
btn_Save_MRA_TemplateToUseForCaseType.Click += MRA_Save_Default_For_CaseType
btn_Save_MRA_TemplateToUseForDept = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Save_MRA_TemplateToUseForDept')
btn_Save_MRA_TemplateToUseForDept.Click += MRA_Save_Default_For_Department

# Edit Questions Area
#lbl_EditRiskAssessment_Name = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_EditRiskAssessment_Name')
#lbl_EditRiskAssessment_ID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_EditRiskAssessment_ID')

dg_MRA_Questions = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_MRA_Questions')
dg_MRA_Questions.SelectionChanged += MRA_Questions_SelectionChanged

#! New 'Publish' button added 21/08/2025
btn_PublishNMRA = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_PublishNMRA')
btn_PublishNMRA.Click += Publish_MRA
tb_ThisMRAid = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_ThisMRAid')
tb_ThisMRAname = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_ThisMRAname')
tb_CopyOfMRAid = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_CopyOfMRAid')
tb_CopyOfMRAname = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_CopyOfMRAname')

btn_AddNew_Q = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_AddNew_Q')
btn_AddNew_Q.Click += AddNew_MRA_Question
btn_CopySelected_Q = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_CopySelected_Q')
btn_CopySelected_Q.Click += Duplicate_MRA_Question
btn_Edit_Q = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Edit_Q')
btn_Q_MoveTop = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Q_MoveTop')
btn_Q_MoveTop.Click += MoveTop_MRA_Question
btn_Q_MoveUp = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Q_MoveUp')
btn_Q_MoveUp.Click += MoveUp_MRA_Question
btn_Q_MoveDown = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Q_MoveDown')
btn_Q_MoveDown.Click += MoveDown_MRA_Question
btn_Q_MoveBottom = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Q_MoveBottom')
btn_Q_MoveBottom.Click += MoveBottom_MRA_Question
btn_DeleteSelected_Q = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_DeleteSelected_Q')
btn_DeleteSelected_Q.Click += Delete_MRA_Question
btn_BackToOverview = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_BackToOverview')
btn_BackToOverview.Click += BackToOverview_MRA_Question
tb_NoQuestions_MRA = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_NoQuestions_MRA')


# Editing Questions - New split view (added 26th April 2024)
grd_EditQs = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'grd_EditQs')
lbl_QID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_QID')
btn_SaveQuestion = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_SaveQuestion')
btn_SaveQuestion.Click += SaveChanges_MRA_Question
txt_QuestionText = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'txt_QuestionText')
cbo_QuestionGroup = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cbo_QuestionGroup')
cbo_QuestionAnswerList = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cbo_QuestionAnswerList')
opt_CopyAnswersFrom = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'opt_CopyAnswersFrom')
opt_BlankAnswerList = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'opt_BlankAnswerList')
btn_AnswerListTypeUpdate = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_AnswerListTypeUpdate')
btn_AnswerListTypeUpdate.Click += updateAnswerList_forSelectedQuestion


# New Editable Answer List (as each Q is now having its own dedicated answer... no longer using 'groups' now we've added 'Email Comment' (which is specific to Question!)
dg_EditMRA_AnswersPreview = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_EditMRA_AnswersPreview')
dg_EditMRA_AnswersPreview.SelectionChanged += dg_EditMRA_AnswersPreview_SelectionChanged
dg_EditMRA_AnswersPreview.CellEditEnding += dg_EditMRA_AnswersPreview_CellEditEnding
lbl_MRA_Answer_Text = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRA_Answer_Text')
lbl_MRA_Answer_Score = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRA_Answer_Score')
lbl_MRA_Answer_EmailComment = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRA_Answer_EmailComment')

btn_AddNewListItem1 = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_AddNewListItem1')
btn_AddNewListItem1.Click += dg_EditMRA_AnswersPreview_addNew
btn_CopySelectedListItem1 = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_CopySelectedListItem1')
btn_CopySelectedListItem1.Click += dg_EditMRA_AnswersPreview_duplicate
btn_A_MoveTop1 = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_A_MoveTop1')
btn_A_MoveTop1.Click += dg_EditMRA_AnswersPreview_moveToTop
btn_A_MoveUp1 = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_A_MoveUp1')
btn_A_MoveUp1.Click += dg_EditMRA_AnswersPreview_moveUp
btn_A_MoveDown1 = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_A_MoveDown1')
btn_A_MoveDown1.Click += dg_EditMRA_AnswersPreview_moveDown
btn_A_MoveBottom1 = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_A_MoveBottom1')
btn_A_MoveBottom1.Click += dg_EditMRA_AnswersPreview_moveToBottom
btn_DeleteSelectedListItem1 = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_DeleteSelectedListItem1')
btn_DeleteSelectedListItem1.Click += dg_EditMRA_AnswersPreview_deleteSelected
lbl_NoAnswers = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_NoAnswers')

# separators on Editing Questions of MRA (in Answers List)
MRA_A_Sep1 = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'MRA_A_Sep1')
MRA_A_Sep2 = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'MRA_A_Sep2')
MRA_A_Sep3 = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'MRA_A_Sep3')
MRA_A_Sep4 = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'MRA_A_Sep4')
MRA_A_Sep5 = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'MRA_A_Sep5')
MRA_A_Sep6 = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'MRA_A_Sep6')


## P R E V I E W   M A T T E R   R I S K   A S S E S S M E N T   - TAB ##
btn_MRAPreview_BackToOverview = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_MRAPreview_BackToOverview')
btn_MRAPreview_BackToOverview.Click += PreviewMRA_BackToOverview
lbl_MRA_Preview_ID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRA_Preview_ID')
lbl_MRA_Preview_Name = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRA_Preview_Name')
lbl_MRAPreview_Score = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRAPreview_Score')
lbl_MRAPreview_RiskCategory = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRAPreview_RiskCategory')
lbl_MRAPreview_RiskCategoryID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRAPreview_RiskCategoryID')


tb_NoMRA_PreviewQs = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_NoMRA_PreviewQs')
dg_MRAPreview = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_MRAPreview')
dg_MRAPreview.SelectionChanged += MRA_Preview_SelectionChanged
dg_GroupItems_Preview = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_GroupItems_Preview')
dg_GroupItems_Preview.SelectionChanged += GroupItems_Preview_SelectionChanged
grid_Preview_MRA = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'grid_Preview_MRA')
lbl_MRAPreview_DGID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRAPreview_DGID')
lbl_MRAPreview_CurrVal = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRAPreview_CurrVal')
chk_MRAPreview_AutoSelectNext = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'chk_MRAPreview_AutoSelectNext')

tb_previewMRA_QestionText = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_previewMRA_QestionText')
cbo_preview_MRA_SelectedComboAnswer = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cbo_preview_MRA_SelectedComboAnswer')
cbo_preview_MRA_SelectedComboAnswer.SelectionChanged += update_EmailComment
tb_preview_MRA_SelectedTextAnswer = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_preview_MRA_SelectedTextAnswer')
tb_preview_MRA_SelectedTextAnswer.TextChanged += update_EmailComment
btn_preview_MRA_SaveAnswer = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_preview_MRA_SaveAnswer')
btn_preview_MRA_SaveAnswer.Click += preview_MRA_SaveAnswer
tb_MRAPreview_EC = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MRAPreview_EC')


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
  </RiskPractice>

</tfb>