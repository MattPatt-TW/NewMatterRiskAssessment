<?xml version="1.0" encoding="UTF-8"?>
<tfb>
  <RiskPracticeV2>
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
UNSELECTED = -1


# # # #   O N   L O A D   E V E N T   # # # #
def myOnLoadEvent(s, event):
  # populate drop-downs
  populate_MRATemplate_CaseTypeCombo()
  
  #MessageBox.Show("DEBUG - POPULATING LISTS")
  refresh_MRA_Templates()                      # Matter Risk Assessment (main overview / templates) 
  
  # hide 'Edit Questions' and 'Preview' tabs
  ti_MRA_Questions.Visibility = Visibility.Collapsed
  ti_MRA_Preview.Visibility = Visibility.Collapsed
  
  # hide other controls not needed until something is selected...
  #tb_ST_NoMRA_Selected.Visibility = Visibility.Visible
  #MessageBox.Show("Hello world! - OnLoad Finished")
  return

# # # #   M A T T E R   R I S K   A S S E S S M E N T   T E M P L A T E S   # # # #

class MRA_Templates(object):
  def __init__(self, myTemplateID, myName, myQCount, myExpiryDays, myRowID, myScoreMedFrom=0, myScoreHighFrom=0):
    self.mraT_TemplateID = myTemplateID
    self.mraT_TemplateName = myName
    self.mraT_QCount = myQCount
    self.mraT_TemplateExpiryDays = myExpiryDays
    self.mraT_ID = myRowID
    self.mraT_ScoreMedFrom = myScoreMedFrom
    self.mraT_ScoreHighFrom = myScoreHighFrom
    return

  def __getitem__(self, index):
    if index == 'TemplateID':           # Note: this is the 'TemplateID' (main ID we'll use whenever linking to other tables) 
      return self.mraT_TemplateID
    elif index == 'ID':                 # Note: this is the actual unique 'ID' (row ID)
      return self.mraT_ID
    elif index == 'Name':
      return self.mraT_TemplateName
    elif index == 'QCount':
      return self.mraT_QCount
    elif index == 'ExpiryDays':
      return self.mraT_TemplateExpiryDays
    else:
      return ''
  

def refresh_MRA_Templates():
  # This funtion populates the main Matter Risk Assessment data grid (and also populates the combo drop-downs in the 'Department' and 'Case Type' defaults area)
  
  # SQL to populate datagrid
  getTableSQL = """SELECT TD.TemplateID, TD.ID, TD.Name, TD.DaysUntil_IncompleteLock, 
                          TD.ScoreMediumTrigger, TD.ScoreHighTrigger, 
                          'Q Count' = (SELECT COUNT(QuestionID) FROM Usr_MRAv2_Templates MRAT WHERE MRAT.TemplateID = TD.TemplateID)
                   FROM Usr_MRAv2_TemplateDetails TD ORDER BY TD.TemplateID"""
  
  tmpItem = []
  
  _tikitDbAccess.Open(getTableSQL)
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          tTemplateID = 0 if dr.IsDBNull(0) else dr.GetValue(0)               
          tID = 0 if dr.IsDBNull(1) else dr.GetValue(1)
          templateName = '' if dr.IsDBNull(2) else dr.GetString(2)
          templateExpiryDays = 0 if dr.IsDBNull(3) else dr.GetValue(3)
          tScoreMedFrom = 0 if dr.IsDBNull(4) else dr.GetValue(4)
          tScoreHighFrom = 0 if dr.IsDBNull(5) else dr.GetValue(5)
          tQCount = 0 if dr.IsDBNull(6) else dr.GetValue(6)
          
          tmpItem.append(MRA_Templates(myTemplateID=tTemplateID, myName=templateName, myQCount=tQCount, 
                                       myExpiryDays=templateExpiryDays, myRowID=tID, myScoreMedFrom=tScoreMedFrom, myScoreHighFrom=tScoreHighFrom))
    dr.Close()
  _tikitDbAccess.Close()
  dg_MRA_Templates.ItemsSource = tmpItem
  return


def dg_MRA_Template_SelectionChanged(s, event):
  # This function will populate the label controls to temp store ID and Name
  #! updated 26/08/2025 to include the other new fields we added

  if dg_MRA_Templates.SelectedIndex == UNSELECTED:
    MRATemplateDetails_ClearFields()
  else:
    MRATemplateDetails_PopulateFieldsFromDataGrid()
  return


def MRATemplateDetails_PopulateFieldsFromDataGrid():
  # This function will populate the fields on the 'Template Details' area for the selected MRA Template

  try:
    selItem = dg_MRA_Templates.SelectedItem
  except:
    #MessageBox.Show("Error obtaining selected Matter Risk Assessment Template details", "Error: Populate Selected Matter Risk Assessment Template Details...")
    return
  
  tb_MRATemplate_Name.Text = str(selItem['Name'])
  lbl_MRATemplate_ID.Content = str(selItem['Code'])
  tb_MRATemplate_ExpiresInXdays.Text = str(selItem['ExpiryDays']) if selItem['ExpiryDays'] is not None else '0'
  tb_ScoreMedFrom.Text = str(selItem['ScoreMedFrom']) if selItem['ScoreMedFrom'] is not None else '0'
  tb_ScoreHighFrom.Text = str(selItem['ScoreHighFrom']) if selItem['ScoreHighFrom'] is not None else '0'
  # now calculate other values
  tb_ScoreLowTo.Text = str(int(selItem['ScoreMedFrom']) - 1) if selItem['ScoreMedFrom'] is not None else '0'
  tb_ScoreMedTo.Text = str(int(selItem['ScoreHighFrom']) - 1) if selItem['ScoreHighFrom'] is not None else '0'


  # set button states
  btn_MRATemplate_CopySelected.IsEnabled = True
  btn_MRATemplate_Preview.IsEnabled = True
  btn_MRATemplate_Edit.IsEnabled = True
  btn_MRATemplate_DeleteSelected.IsEnabled = True if int(dg_MRA_Templates.SelectedItem['CountUsed']) == 0 else False
  btn_MRATemplate_SaveHeaderDetails.IsEnabled = True

  # load case type defaults for this template
  refresh_MRA_CaseType_Defaults()
  return


def MRATemplateDetails_ClearFields():
  # This function will clear the fields on the 'Template Details' area for the selected MRA Template

  tb_MRATemplate_Name.Text = ''
  lbl_MRATemplate_ID.Content = ''
  tb_MRATemplate_ExpiresInXdays.Text = '0'
  tb_ScoreLowTo.Text = '0'
  tb_ScoreMedFrom.Text = '0'
  tb_ScoreMedTo.Text = '0'
  tb_ScoreHighFrom.Text = '0'

  # set button states
  btn_MRATemplate_CopySelected.IsEnabled = False
  btn_MRATemplate_Preview.IsEnabled = False
  btn_MRATemplate_Edit.IsEnabled = False
  btn_MRATemplate_DeleteSelected.IsEnabled = False
  btn_MRATemplate_SaveHeaderDetails.IsEnabled = False 

  # clear case type defaults datagrid
  dg_MRATemplate_CaseTypes.ItemsSource = []
  return


def btn_MRATemplate_SaveHeaderDetails_Click(s, event):
  # This is the 'Save' button on the 'List of NMRA Templates' tab, and saves the 'header'/details to the selected template

  # firstly check something is selected
  if dg_MRA_Templates.SelectedIndex == -1:
    MessageBox.Show("No Matter Risk Assessment Template has been selected!\nPlease select a template before clicking 'Save Changes'", "Error: Save Changes to Selected Matter Risk Assessment Template...")
    return

  itemID = lbl_MRATemplate_ID.Content
  newName = str(tb_MRATemplate_Name.Text)
  newName = newName.replace("'", "''")

  try:
    newExpDays = int(tb_MRATemplate_ExpiresInXdays.Text)
  except:
    MessageBox.Show("The 'Expires in ?? days' value must be a whole number (integer)", "Error: Save Changes to Selected NMRA Template...")
    return
  
  try:
    newScoreMedFrom = int(tb_ScoreMedFrom.Text)
    newScoreHighFrom = int(tb_ScoreHighFrom.Text)
  except:
    MessageBox.Show("The 'Score Medium From' and 'Score High From' values must be whole numbers (integers)", "Error: Save Changes to Selected NMRA Template...")
    return
  
  # form the SQL to update
  updateSQL = """UPDATE Usr_MRAv2_TemplateDetails SET Name = '{0}', DaysUntil_IncompleteLock = {1}, ScoreMediumTrigger = {2}, ScoreHighTrigger = {3}
                 WHERE TemplateID = {4}""".format(newName, newExpDays, newScoreMedFrom, newScoreHighFrom, itemID)

  # do update
  try:
    _tikitResolver.Resolve("[SQL: {0}]".format(updateSQL))
    
    MessageBox.Show("Successfully updated details of selected Matter Risk Assessment Template", "Save Changes to Selected NMRA Template - Success...")
  except:
    MessageBox.Show("There was an error amending the details of the Matter Risk Assessment Template, using SQL:\n" + str(updateSQL), "Error: Amending Details of Matter Risk Assessment Template...")
    return
  
  MRATemplates_refreshAndReselect(withTemplateID=itemID)
  return


def btn_MRATemplate_AddNew_Click(s, event):
  # This function will add a new row to the 'MRAv2_TemplateDetails' table
  
  #! Added 29/07/2025: Get next new TypeID so we can add it directly in the INSERT statement
  nextTypeID = runSQL(codeToRun="SELECT ISNULL(MAX(TemplateID), 0) + 1 FROM Usr_MRAv2_TemplateDetails", returnType='Int')
  insertSQL = """[SQL: INSERT INTO Usr_MRAv2_TemplateDetails (Name, DaysUntil_IncompleteLock, ScoreMediumTrigger, ScoreHighTrigger, TemplateID)
                 VALUES ('NMRA - new', 29, 0, 0, {0})]""".format(nextTypeID)
  
  try:
    _tikitResolver.Resolve(insertSQL)
  except:
    MessageBox.Show("There was an error trying to create a new Matter Risk Assessment, using SQL:\n" + str(insertSQL), "Error: Adding new Matter Risk Assessment...")
    return
  
  # refresh data grid and select last item
  MRATemplates_refreshAndReselect(withTemplateID=nextTypeID)
  return


def MRATemplates_refreshAndReselect(withTemplateID=None):
  # This function will refresh the MRA Templates data grid and re-select the current selected item (if possible)
  # get current itemID
  if withTemplateID is not None:
    currentSelectedID = withTemplateID
  else:
    currentSelectedID = dg_MRA_Templates.SelectedItem['TemplateID'] if dg_MRA_Templates.SelectedIndex != -1 else None

  refresh_MRA_Templates()

  # reselect item
  if currentSelectedID is not None:
    tCount = -1
    for x in dg_MRA_Templates.Items:
      tCount += 1
      if str(x['TemplateID']) == str(currentSelectedID):
        dg_MRA_Templates.SelectedIndex = tCount
        dg_MRA_Templates.ScrollIntoView(dg_MRA_Templates.Items[tCount])
        break

  return


def btn_MRATemplate_CopySelected_Click(s, event):
  # This function will duplicate the selected Matter Risk Assessment (including the questions) AND ANSWERS - NEED TO REVIEW

  MessageBox.Show("Copy selected MRA button click", "Duplicating Matter Risk Assessment...")
  return


def btn_MRATemplate_DeleteSelected_Click(s, event):
  # This function will delete the selected Matter Risk Assessment template (and any questions associated to it)

  MessageBox.Show("Delete selected MRA button click", "Delete Matter Risk Assessment...")
  return
  
  
def btn_MRATemplate_Preview_Click(s, event):
  # This function will load the 'Preview' tab (made to look like 'matter-level' XAML) for the selected item

  MessageBox.Show("Preview selected MRA click", "Preview selected Matter Risk Assessment...")
  return
  
  
def btn_MRATemplate_Edit_Click(s, event):
  # This function will load the 'Questions' tab for the selected item
  
  # if nothing selected, alert user and bomb-out now...
  if dg_MRA_Templates.SelectedIndex == -1:
    MessageBox.Show("Nothing selected to Edit!", "Error: Edit selected Matter Risk Assessment...")
    return
  else:
    origItem = dg_MRA_Templates.SelectedItem  
  
  MessageBox.Show("EditSelected_Click", "DEBUG - TESTING")
  return

# # # #  END OF:  Matter Risk Assessment Templates   # # # #

 
# # # #   C A S E   T Y P E   D E F A U L T S   # # # #
class caseType_Defaults(object):
  def __init__(self, myCaseTypeName, myRowID, myCaseTypeID, myDeptName, myDeptID):
    self.mraT_CaseTypeName = myCaseTypeName 
    self.mraT_RowID = myRowID
    self.mraT_CaseTypeID = myCaseTypeID
    self.mraT_DeptName = myDeptName
    self.mraT_DeptID = myDeptID
    return
    
  def __getitem__(self, index):
    if index == 'CaseTypeName':
      return self.mraT_CaseTypeName
    elif index == 'CaseTypeID':
      return self.mraT_CaseTypeID
    elif index == 'DeptName':
      return self.mraT_DeptName
    elif index == 'DeptID':
      return self.mraT_DeptID
    elif index == 'RowID':
      return self.mraT_RowID
    else:
      return ''
      
def refresh_MRA_CaseType_Defaults():
  # This function will populate the 'Case Types' datagrid (for selecting which Matter Risk Assessment template to be used)
  
  # if nothing selected, exit now
  if dg_MRA_Templates.SelectedIndex == UNSELECTED:
    dg_MRATemplate_CaseTypes.ItemsSource = []
    return

  getTableSQL = """SELECT '0-RowID' = CTD.ID, 
                          '1-CaseTypeName' = CT.Description, 
                          '2-CaseTypeID' = CTD.CaseTypesCode,
                          '3-DeptName' = CTG.Name,
                          '4-DeptID' = CT.CaseTypeGroupRef
                   FROM Usr_MRAv2_CaseTypeDefaults CTD
                      LEFT OUTER JOIN Usr_MRAv2_TemplateDetails TD ON CTD.TemplateID = TD.TemplateID
                      LEFT OUTER JOIN CaseTypes CT ON CTD.CaseTypesCode = CT.Code
                      LEFT OUTER JOIN CaseTypeGroups CTG ON CT.CaseTypeGroupRef = CTG.ID
                   WHERE CTD.TemplateID = {templateID}""".format(templateID=dg_MRA_Templates.SelectedItem['TemplateID'])

  tmpItem = []
  _tikitDbAccess.Open(getTableSQL)
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          tmpRowID = 0 if dr.IsDBNull(0) else dr.GetValue(0)
          tmpCaseTypeName = '' if dr.IsDBNull(1) else dr.GetString(1) 
          tmpCaseTypeID = 0 if dr.IsDBNull(2) else dr.GetValue(2)
          tmpDeptName = '' if dr.IsDBNull(3) else dr.GetString(3)
          tmpDeptID = 0 if dr.IsDBNull(4) else dr.GetValue(4)
          
          tmpItem.append(caseType_Defaults(tmpCaseTypeName, tmpRowID, tmpCaseTypeID, tmpDeptName, tmpDeptID))

    dr.Close()
  
  #close db connection
  _tikitDbAccess.Close()
  
  # added the following 19th June 2024 - this will display list grouped by 'Department' (CaseTypeGroup) name
  # note: added ', CollectionView, ListCollectionView, PropertyGroupDescription' to 'from System.Windows.Data import Binding ' (line 20)
  tmpC = ListCollectionView(tmpItem)
  tmpC.GroupDescriptions.Add(PropertyGroupDescription("mraT_DeptName"))

  dg_MRATemplate_CaseTypes.ItemsSource = tmpC     #tmpItem    # use 'tmpItem' if don't want to show 'groupings' in DataGrid
  return


def btn_CaseTypeLinkToTemplate_add_Click(s, event):
  # This function will link the selected Case Type to the current selected Matter Risk Assessment Template
  
  MessageBox.Show("Add Case Type to Template link button click", "Link Case Type to Matter Risk Assessment Template...")
  return

def btn_CaseTypeLinkToTemplate_remove_Click(s, event):
  # This function will remove the link between the selected Case Type and the current selected Matter Risk Assessment Template
  
  MessageBox.Show("Remove Case Type to Template link button click", "Unlink Case Type from Matter Risk Assessment Template...")
  return


  
def btn_BackToOverview_FromEditQs_Click(s, event):
  # This function should clear the 'Questions' tab and take us back to the 'Risk Assessment Overview' tab
  ti_MRA_Questions.Visibility = Visibility.Collapsed
  ti_MRA_Overview.Visibility = Visibility.Visible
  ti_MRA_Overview.IsSelected = True
  refresh_MRA_Templates()
  return
  
  
class caseType_Item(object):
  def __init__(self, myCaseTypeID, myCaseTypeName, myDeptName):
    self.mraT_CaseTypeCode = myCaseTypeID
    self.mraT_CaseTypeName = myCaseTypeName
    self.mraT_DepartmentName = myDeptName
    return
    
  def __getitem__(self, index):
    if index == 'CaseTypeID':
      return self.mraT_CaseTypeCode
    elif index == 'CaseTypeName':
      return self.mraT_CaseTypeName
    elif index == 'DeptName':
      return self.mraT_DepartmentName
    else:
      return ''
    
def populate_MRATemplate_CaseTypeCombo():
  # This function will populate the 'Case Type' combo box on the 'Template Details' area
  
  getSQL = """SELECT 'Dept' = CTG.Name, 'CaseType Desc' = CT.Description, 'CaseTypeCode' = CT.Code 
              FROM CaseTypes CT 
                LEFT OUTER JOIN CaseTypeGroups CTG ON CT.CaseTypeGroupRef = CTG.ID
              ORDER BY CTG.Name, CT.Description"""

  tmpItem = []
  _tikitDbAccess.Open(getSQL)
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          tmpDeptName = '' if dr.IsDBNull(0) else dr.GetString(0)
          tmpCaseTypeName = '' if dr.IsDBNull(1) else dr.GetString(1)
          tmpCaseTypeID = 0 if dr.IsDBNull(2) else dr.GetValue(2)
          
          tmpItem.append(caseType_Item(myCaseTypeID=tmpCaseTypeID, myCaseTypeName=tmpCaseTypeName, myDeptName=tmpDeptName))

    dr.Close()
  _tikitDbAccess.Close()

  tmpList = ListCollectionView(tmpItem)
  # group on Department name
  tmpList.GroupDescriptions.Add(PropertyGroupDescription("mraT_DepartmentName"))
  cbo_MRATemplate_CaseType.ItemsSource = tmpList
  return



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
#TC_Main = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'TC_Main')
ti_MRA_Overview = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_MRA_Overview')
ti_MRA_Questions = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_MRA_Questions')
ti_MRA_Preview = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_MRA_Preview')


## C O N F I G U R E   M A T T E R   R I S K   A S S E S S M E N T    - TAB ##
dg_MRA_Templates = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_MRA_Templates')
dg_MRA_Templates.SelectionChanged += dg_MRA_Template_SelectionChanged

#! Edit MRA Template Details
tb_MRATemplate_Name = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MRATemplate_Name')            # TextBox for editing the name of the selected NMRA template 
lbl_MRATemplate_ID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRATemplate_ID')              # Label to store ID of selected NMRA template
tb_MRATemplate_ExpiresInXdays = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MRATemplate_ExpiresInXdays')

tb_ScoreLowTo = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_ScoreLowTo')
tb_ScoreMedFrom = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_ScoreMedFrom')
tb_ScoreMedTo = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_ScoreMedTo')
tb_ScoreHighFrom = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_ScoreHighFrom')

#! Buttons for MRA Template management
btn_MRATemplate_AddNew = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_MRATemplate_AddNew')
btn_MRATemplate_AddNew.Click += btn_MRATemplate_AddNew_Click

btn_MRATemplate_SaveHeaderDetails = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_MRATemplate_SaveHeaderDetails')
btn_MRATemplate_SaveHeaderDetails.Click += btn_MRATemplate_SaveHeaderDetails_Click

# still to do below
btn_MRATemplate_CopySelected = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_MRATemplate_CopySelected')
btn_MRATemplate_CopySelected.Click += btn_MRATemplate_CopySelected_Click
btn_MRATemplate_Preview = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_MRATemplate_Preview')
btn_MRATemplate_Preview.Click += btn_MRATemplate_Preview_Click
btn_MRATemplate_Edit = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_MRATemplate_Edit')
btn_MRATemplate_Edit.Click += btn_MRATemplate_Edit_Click
btn_MRATemplate_DeleteSelected = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_MRATemplate_DeleteSelected')
btn_MRATemplate_DeleteSelected.Click += btn_MRATemplate_DeleteSelected_Click

# MRA Template to Case Type Linking Area
dg_MRATemplate_CaseTypes = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_MRATemplate_CaseTypes')
btn_CaseTypeLinkToTemplate_add = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_CaseTypeLinkToTemplate_add')
btn_CaseTypeLinkToTemplate_add.Click += btn_CaseTypeLinkToTemplate_add_Click
btn_CaseTypeLinkToTemplate_remove = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_CaseTypeLinkToTemplate_remove')
btn_CaseTypeLinkToTemplate_remove.Click += btn_CaseTypeLinkToTemplate_remove_Click
cbo_MRATemplate_CaseType = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cbo_MRATemplate_CaseType')

#########################################################################################


# Edit Questions Area
#lbl_EditRiskAssessment_Name = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_EditRiskAssessment_Name')
#lbl_EditRiskAssessment_ID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_EditRiskAssessment_ID')

dg_MRA_Questions = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_MRA_Questions')
#dg_MRA_Questions.SelectionChanged += MRA_Questions_SelectionChanged

#! New 'Publish' button added 21/08/2025
tb_ThisMRAid = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_ThisMRAid')
tb_ThisMRAname = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_ThisMRAname')
tb_CopyOfMRAid = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_CopyOfMRAid')
tb_CopyOfMRAname = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_CopyOfMRAname')

btn_AddNew_Q = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_AddNew_Q')
#btn_AddNew_Q.Click += AddNew_MRA_Question
btn_CopySelected_Q = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_CopySelected_Q')
#btn_CopySelected_Q.Click += Duplicate_MRA_Question
btn_Edit_Q = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Edit_Q')
btn_Q_MoveTop = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Q_MoveTop')
#btn_Q_MoveTop.Click += MoveTop_MRA_Question
btn_Q_MoveUp = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Q_MoveUp')
#btn_Q_MoveUp.Click += MoveUp_MRA_Question
btn_Q_MoveDown = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Q_MoveDown')
#btn_Q_MoveDown.Click += MoveDown_MRA_Question
btn_Q_MoveBottom = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Q_MoveBottom')
#btn_Q_MoveBottom.Click += MoveBottom_MRA_Question
btn_DeleteSelected_Q = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_DeleteSelected_Q')
#btn_DeleteSelected_Q.Click += Delete_MRA_Question
btn_BackToOverview_FromEditQs = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_BackToOverview_FromEditQs')
btn_BackToOverview_FromEditQs.Click += btn_BackToOverview_FromEditQs_Click
tb_NoQuestions_MRA = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_NoQuestions_MRA')


# Editing Questions - New split view (added 26th April 2024)
grd_EditQs = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'grd_EditQs')
lbl_QID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_QID')
btn_SaveQuestion = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_SaveQuestion')
#btn_SaveQuestion.Click += SaveChanges_MRA_Question
txt_QuestionText = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'txt_QuestionText')
cbo_QuestionGroup = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cbo_QuestionGroup')
cbo_QuestionAnswerList = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cbo_QuestionAnswerList')
opt_CopyAnswersFrom = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'opt_CopyAnswersFrom')
opt_BlankAnswerList = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'opt_BlankAnswerList')
btn_AnswerListTypeUpdate = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_AnswerListTypeUpdate')
#btn_AnswerListTypeUpdate.Click += updateAnswerList_forSelectedQuestion


# New Editable Answer List (as each Q is now having its own dedicated answer... no longer using 'groups' now we've added 'Email Comment' (which is specific to Question!)
dg_EditMRA_AnswersPreview = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_EditMRA_AnswersPreview')
#dg_EditMRA_AnswersPreview.SelectionChanged += dg_EditMRA_AnswersPreview_SelectionChanged
#dg_EditMRA_AnswersPreview.CellEditEnding += dg_EditMRA_AnswersPreview_CellEditEnding
lbl_MRA_Answer_Text = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRA_Answer_Text')
lbl_MRA_Answer_Score = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRA_Answer_Score')
lbl_MRA_Answer_EmailComment = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRA_Answer_EmailComment')

btn_AddNewListItem1 = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_AddNewListItem1')
#btn_AddNewListItem1.Click += dg_EditMRA_AnswersPreview_addNew
btn_CopySelectedListItem1 = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_CopySelectedListItem1')
#btn_CopySelectedListItem1.Click += dg_EditMRA_AnswersPreview_duplicate
btn_A_MoveTop1 = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_A_MoveTop1')
#btn_A_MoveTop1.Click += dg_EditMRA_AnswersPreview_moveToTop
btn_A_MoveUp1 = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_A_MoveUp1')
#btn_A_MoveUp1.Click += dg_EditMRA_AnswersPreview_moveUp
btn_A_MoveDown1 = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_A_MoveDown1')
#btn_A_MoveDown1.Click += dg_EditMRA_AnswersPreview_moveDown
btn_A_MoveBottom1 = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_A_MoveBottom1')
#btn_A_MoveBottom1.Click += dg_EditMRA_AnswersPreview_moveToBottom
btn_DeleteSelectedListItem1 = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_DeleteSelectedListItem1')
#btn_DeleteSelectedListItem1.Click += dg_EditMRA_AnswersPreview_deleteSelected
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
#dg_MRAPreview.SelectionChanged += MRA_Preview_SelectionChanged
dg_GroupItems_Preview = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_GroupItems_Preview')
#dg_GroupItems_Preview.SelectionChanged += GroupItems_Preview_SelectionChanged
grid_Preview_MRA = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'grid_Preview_MRA')
lbl_MRAPreview_DGID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRAPreview_DGID')
lbl_MRAPreview_CurrVal = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRAPreview_CurrVal')
chk_MRAPreview_AutoSelectNext = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'chk_MRAPreview_AutoSelectNext')

tb_previewMRA_QestionText = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_previewMRA_QestionText')
cbo_preview_MRA_SelectedComboAnswer = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cbo_preview_MRA_SelectedComboAnswer')
#cbo_preview_MRA_SelectedComboAnswer.SelectionChanged += update_EmailComment
tb_preview_MRA_SelectedTextAnswer = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_preview_MRA_SelectedTextAnswer')
#tb_preview_MRA_SelectedTextAnswer.TextChanged += update_EmailComment
btn_preview_MRA_SaveAnswer = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_preview_MRA_SaveAnswer')
#btn_preview_MRA_SaveAnswer.Click += preview_MRA_SaveAnswer
tb_MRAPreview_EC = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MRAPreview_EC')




# Define Actions and on load events
myOnLoadEvent(_tikitSender, 'onLoad')

]]>
    </Loaded>
  </RiskPracticeV2>

</tfb>