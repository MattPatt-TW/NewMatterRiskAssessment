<?xml version="1.0" encoding="UTF-8"?>
<tfb>
  <RiskPracticeV2>
    <Init>
      <![CDATA[
import clr
#from TWUtils import runSQL
clr.AddReference("System")            # for new MRA Edit Template tab code
clr.AddReference("WindowsBase")       # for new MRA Edit Template tab code

clr.AddReference('mscorlib')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('System.Windows.Forms')

from datetime import datetime
from System import DateTime, EventHandler
from System.Diagnostics import Process
from System.Globalization import DateTimeStyles
from System.Collections.Generic import Dictionary
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Collections.ObjectModel import ObservableCollection
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
_temp_id = -1

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
    elif index == 'ScoreMedFrom':
      return self.mraT_ScoreMedFrom
    elif index == 'ScoreHighFrom':
      return self.mraT_ScoreHighFrom
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
  dg_MRA_Templates_SetVisibilityOfEditArea()
  return

def dg_MRA_Templates_SetVisibilityOfEditArea():
  # This function will set the visibility of the 'Template Details' edit area depending on whether something is selected in the datagrid or not

  # if nothing in list, hide 'edit' fields and exit early
  if dg_MRA_Templates.Items.Count == 0:
    #dg_MRA_Templates.Visibility = Visibility.Collapsed
    tb_MRA_NoneSelected.Visibility = Visibility.Visible
    stk_ST_SelectedMRA.Visibility = Visibility.Collapsed
    return

  # set visibility based on selection
  if dg_MRA_Templates.SelectedIndex == UNSELECTED:
    tb_MRA_NoneSelected.Visibility = Visibility.Visible
    stk_ST_SelectedMRA.Visibility = Visibility.Collapsed
  else:
    tb_MRA_NoneSelected.Visibility = Visibility.Collapsed
    stk_ST_SelectedMRA.Visibility = Visibility.Visible
  return

def dg_MRA_Template_SelectionChanged(s, event):
  # This function will populate the label controls to temp store ID and Name
  #! updated 26/08/2025 to include the other new fields we added

  if dg_MRA_Templates.SelectedIndex == UNSELECTED:
    MRATemplateDetails_ClearFields()
  else:
    MRATemplateDetails_PopulateFieldsFromDataGrid()
  dg_MRA_Templates_SetVisibilityOfEditArea()
  return


def MRATemplateDetails_PopulateFieldsFromDataGrid():
  # This function will populate the fields on the 'Template Details' area for the selected MRA Template

  try:
    selItem = dg_MRA_Templates.SelectedItem
  except:
    #MessageBox.Show("Error obtaining selected Matter Risk Assessment Template details", "Error: Populate Selected Matter Risk Assessment Template Details...")
    return
  
  tb_MRATemplate_Name.Text = str(selItem['Name'])
  lbl_MRATemplate_ID.Content = str(selItem['TemplateID'])
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
  btn_MRATemplate_DeleteSelected.IsEnabled = True #if int(dg_MRA_Templates.SelectedItem['CountUsed']) == 0 else False
  btn_MRATemplate_SaveHeaderDetails.IsEnabled = True

  # load case type defaults DataGrid for this template
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
  if lbl_MRATemplate_ID.Content == '0' or dg_MRA_Templates.SelectedIndex == UNSELECTED:
    MessageBox.Show("Nothing selected to Edit!", "Error: Edit selected Matter Risk Assessment...")
    return

  # put details into header area of 'Questions' tab
  tb_ThisMRAid.Text = str(lbl_MRATemplate_ID.Content)
  tb_ThisMRAname.Text = str(tb_MRATemplate_Name.Text)

  # refresh questions datagrid
  #dg_MRA_Questions_Refresh()

  # new treeview/template editor VM code
  EditMRA_loadTreeViewStructure(selectedTemplateID=int(tb_ThisMRAid.Text))

  # show 'Questions' tab and hide 'Overview' tab
  ti_MRA_Overview.Visibility = Visibility.Collapsed
  ti_MRA_Questions.Visibility = Visibility.Visible
  ti_MRA_Questions.IsSelected = True
  #MessageBox.Show("EditSelected_Click", "DEBUG - TESTING")
  return

# # # #  END OF:  Matter Risk Assessment Templates   # # # #

# # # #    Q U E S T I O N S   # # # #
class MRA_Questions(object):
  def __init__(self, myQuestionID, myQuestionText, myQuestionGroup, myDisplayOrder, myCountAnswers):
    self.mraQ_QuestionID = myQuestionID
    self.mraQ_QuestionText = myQuestionText
    self.mraQ_QuestionGroup = myQuestionGroup
    self.mraQ_QDisplayOrder = myDisplayOrder
    self.mraQ_CountAnswers = myCountAnswers
    return

  def __getitem__(self, index):
    if index == 'QuestionID':
      return self.mraQ_QuestionID
    elif index == 'QuestionText':
      return self.mraQ_QuestionText
    elif index == 'QuestionGroup':
      return self.mraQ_QuestionGroup
    elif index == 'DisplayOrder':
      return self.mraQ_QDisplayOrder
    elif index == 'CountAnswers':
      return self.mraQ_CountAnswers
    else:
      return ''


def btn_Questions_AddNew_Click(s, event):
  # This function will add a new question to the current selected Matter Risk Assessment template
  
  templateID = int(tb_ThisMRAid.Text)
  questionID = runSQL(codeToRun="SELECT ISNULL(MAX(QuestionID), 0) + 1 FROM Usr_MRAv2_Question", returnType='Int')
  insertSQL = """INSERT INTO Usr_MRAv2_Question (QuestionID, QuestionText)
                 VALUES ({qID}, 'New Question - please edit text')""".format(qID=questionID)

  MessageBox.Show("Add New Question button click", "Add New Question...")
  return

def btn_Questions_Clipboard_Click(s, event):
  # This function will open the 'Questions Clipboard' window for copying/pasting questions between templates
  
  QuestionsClipboard_Popup.IsOpen = btn_Questions_Clipboard.IsChecked
  #MessageBox.Show("Questions Clipboard button click", "Open Questions Clipboard...")
  return

def QuestionsClipboard_Popup_Closed(s, event):
  # This function will uncheck the 'Questions Clipboard' button when the popup is closed
  
  btn_Questions_Clipboard.IsChecked = False
  QuestionsClipboard_Popup.IsOpen = False
  return

def mi_Question_CopyToClipboard_Click(s, event):
  # This function will copy the selected question to the clipboard
  
  QuestionsClipboard_Popup_Closed(s, event)
  MessageBox.Show("Copy Question to Clipboard menu item click", "Copy Question to Clipboard...")
  return

def mi_Question_PasteFromClipboard_Click(s, event):
  # This function will paste the question from the clipboard to the current template
  
  QuestionsClipboard_Popup_Closed(s, event)
  MessageBox.Show("Paste Question from Clipboard menu item click", "Paste Question from Clipboard...")
  return


def btn_Question_MoveTop_Click(s, event):
  # This function will move the selected question to the top of the list
  
  MessageBox.Show("Move Question to Top button click", "Move Question to Top...")
  return

def btn_Question_MoveUp_Click(s, event):
  # This function will move the selected question up one position in the list
  
  MessageBox.Show("Move Question Up button click", "Move Question Up...")
  return

def btn_Question_MoveDown_Click(s, event):
  # This function will move the selected question down one position in the list
  
  MessageBox.Show("Move Question Down button click", "Move Question Down...")
  return

def btn_Question_MoveBottom_Click(s, event):
  # This function will move the selected question to the bottom of the list
  
  MessageBox.Show("Move Question to Bottom button click", "Move Question to Bottom...")
  return

def btn_Question_DeleteSelected_Click(s, event):
  # This function will delete the selected question from the current template
  
  MessageBox.Show("Delete Selected Question button click", "Delete Selected Question...")
  return

def dg_MRA_Questions_SelectionChanged(s, event):
  # This function will handle when the selection changes in the 'MRA Questions' datagrid - it puts selected question details into the edit area below

  if dg_MRA_Questions.SelectedIndex == UNSELECTED:
    dg_MRA_Questions_EditArea_Clear()
    return
  else:
    dg_MRA_Questions_EditArea_PopulateFromSelected()
  
  # for now, just show a message box
  #MessageBox.Show("Selection changed in MRA Questions datagrid", "DEBUG - MRA Questions Selection Changed...")
  return

def dg_MRA_Questions_EditArea_Clear():
  # This function will clear the 'Edit Question' area below the 'MRA Questions' datagrid
  
  tb_ESQ_QuestionID.Text = ''
  txt_ESQ_QuestionText.Text = ''
  txt_ESQ_QuestionGroup.Text = ''
  return

def dg_MRA_Questions_EditArea_PopulateFromSelected():
  # This function will populate the 'Edit Question' area below the 'MRA Questions' datagrid from the selected question
  
  selItem = dg_MRA_Questions.SelectedItem
  tb_ESQ_QuestionID.Text = str(selItem['QuestionID'])
  txt_ESQ_QuestionText.Text = str(selItem['QuestionText'])
  txt_ESQ_QuestionGroup.Text = str(selItem['QuestionGroup'])
  return


def dg_MRA_Questions_Refresh():
  # This function will refresh the 'MRA Questions' datagrid for the selected template
  
  # otherwise, get TemplateID to use from this page
  templateID = int(tb_ThisMRAid.Text)
  if templateID == -1:
    return

  # form SQL to get questions for this template
  sql = """SELECT '0-QuestionID' = T.QuestionID,
                  '1-QuestionText' = Q.QuestionText,
                  '2-QuestionGroup' = T.QuestionGroup,
                  '3-DisplayOrder' = T.DisplayOrder,
                  '4-CountAs' = (SELECT COUNT(T1.AnswerID) FROM Usr_MRAv2_Templates T1 WHERE T1.TemplateID = {tID} AND T1.QuestionID = T.QuestionID)
            FROM Usr_MRAv2_Templates T
                JOIN Usr_MRAv2_Question Q ON T.QuestionID = Q.QuestionID
            WHERE T.TemplateID = {tID}
            GROUP BY T.QuestionID, Q.QuestionText, T.QuestionGroup, T.DisplayOrder
            ORDER BY T.QuestionGroup, T.DisplayOrder""".format(tID=templateID)

  tmpItem = []
  _tikitDbAccess.Open(sql)
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          tmpQuestionID = 0 if dr.IsDBNull(0) else dr.GetValue(0)
          tmpQuestionText = '' if dr.IsDBNull(1) else dr.GetString(1)
          tmpQuestionGroup = '' if dr.IsDBNull(2) else dr.GetString(2)
          tmpDisplayOrder = 0 if dr.IsDBNull(3) else dr.GetValue(3)
          tmpCountAnswers = 0 if dr.IsDBNull(4) else dr.GetValue(4)
          
          tmpItem.append(MRA_Questions(myQuestionID=tmpQuestionID, myQuestionText=tmpQuestionText, myQuestionGroup=tmpQuestionGroup,
                                       myDisplayOrder=tmpDisplayOrder, myCountAnswers=tmpCountAnswers))
    dr.Close()
  _tikitDbAccess.Close()

  # add Grouping on 'QuestionGroup'
  # note: added ', CollectionView, ListCollectionView, PropertyGroupDescription' to 'from System.Windows.Data import Binding ' (line 20)
  tmpC = ListCollectionView(tmpItem)
  tmpC.GroupDescriptions.Add(PropertyGroupDescription("mraQ_QuestionGroup"))
  dg_MRA_Questions.ItemsSource = tmpC

  dg_MRA_Questions_SetVisibilityOfEditArea()
  #MessageBox.Show("Refresh MRA Questions datagrid", "DEBUG - Refresh MRA Questions...")
  return

def dg_MRA_Questions_SetVisibilityOfEditArea():
  # This function will set the visibility of the 'Edit Question' area below the datagrid depending on whether something is selected or not

  if dg_MRA_Questions.Items.Count == 0:
    tb_MRA_NoQuestionsText.Visibility = Visibility.Visible
    dg_MRA_Questions.Visibility = Visibility.Collapsed
  else:
    tb_MRA_NoQuestionsText.Visibility = Visibility.Collapsed
    dg_MRA_Questions.Visibility = Visibility.Visible

  if dg_MRA_Questions.SelectedIndex == UNSELECTED:
    tb_ESQ_QuestionID.Text = ''
    txt_ESQ_QuestionText.Text = ''
    txt_ESQ_QuestionGroup.Text = ''
    btn_ESQ_SaveQuestion.IsEnabled = False
    tb_ESQ_QuestionID.IsEnabled = False
    txt_ESQ_QuestionText.IsEnabled = False
    txt_ESQ_QuestionGroup.IsEnabled = False
  else:
    btn_ESQ_SaveQuestion.IsEnabled = True
    tb_ESQ_QuestionID.IsEnabled = True
    txt_ESQ_QuestionText.IsEnabled = True
    txt_ESQ_QuestionGroup.IsEnabled = True
  return


def btn_ESQ_SaveQuestion_Click(s, event):
  # This function will save the edited question details from the 'Edit Question' area below the datagrid
  
  #! Note: because of new structure we need to:
  # 1) update 'Usr_MRAv2_Question' table for question text -
  #    first check if this text already exists
  #      - if yes, use THAT QuestionID instead (avoids having multiple identical questions in the table);
  #      - if no, update question text for current QuestionID
  # 2) update 'Usr_MRAv2_Templates' table for question group and display order (using the QuestionID from step 1)
  # Note: we may need to also update any 'Answers' linked to this question if the QuestionID changes (ie: new question text added)
  #       so will want an 'originalQuestionID' variable to use for 'UPDATE' substitutions if needed (as shouldn't just delete/ignore current answers). 

  MessageBox.Show("Save Edited Question button click", "Save Edited Question...")
  return
# # # #   E N D   O F :   Q U E S T I O N S   # # # #


 
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
          
          tmpItem.append(caseType_Defaults(myCaseTypeName=tmpCaseTypeName, myRowID=tmpRowID, 
                                           myCaseTypeID=tmpCaseTypeID, myDeptName=tmpDeptName, myDeptID=tmpDeptID))

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
  
  templateID = lbl_MRATemplate_ID.Content      #dg_MRA_Templates.SelectedItem['TemplateID']
  caseTypeID = cbo_MRATemplate_CaseType.SelectedItem['CaseTypeID']

  # first check if this link already exists
  checkSQL = """SELECT COUNT(*) FROM Usr_MRAv2_CaseTypeDefaults 
                WHERE TemplateID = {tID} AND CaseTypesCode = {ctID}""".format(tID=templateID, ctID=caseTypeID)
  existingCount = runSQL(codeToRun=checkSQL, returnType='Int')

  if existingCount > 0:
    MessageBox.Show("The selected Case Type is already linked to this Matter Risk Assessment Template!", "Error: Link Case Type to Matter Risk Assessment Template...")
    return
  
  # else add the link
  insertSQL = """INSERT INTO Usr_MRAv2_CaseTypeDefaults (TemplateID, CaseTypesCode)
                 VALUES ({tID}, {ctID})""".format(tID=templateID, ctID=caseTypeID)
  try:
    _tikitResolver.Resolve("[SQL: {0}]".format(insertSQL))
    refresh_MRA_CaseType_Defaults()
  except:
    MessageBox.Show("There was an error trying to link the selected Case Type to this Matter Risk Assessment Template, using SQL:\n" + str(insertSQL), "Error: Link Case Type to Matter Risk Assessment Template...")
    return
  
  #MessageBox.Show("Add Case Type to Template link button click", "Link Case Type to Matter Risk Assessment Template...")
  return

def btn_CaseTypeLinkToTemplate_remove_Click(s, event):
  # This function will remove the link between the selected Case Type and the current selected Matter Risk Assessment Template

  if dg_MRATemplate_CaseTypes.SelectedIndex == UNSELECTED:
    MessageBox.Show("No Case Type selected to remove the link for!", "Error: Unlink Case Type from Matter Risk Assessment Template...")
    return
  
  # else, double-check user wants to do this - if no, exit now
  confirmResult = MessageBox.Show("Are you sure you want to unlink the selected Case Type from this Matter Risk Assessment Template?", 
                                  "Confirm: Unlink Case Type from Matter Risk Assessment Template...", 
                                  MessageBoxButtons.YesNo, MessageBoxIcon.Question)
  if confirmResult == DialogResult.No:
    return

  # continue to remove link
  # get IDs to use - making sure to convert to integers to avoid SQL errors (exit early if can't be converted to integer)
  try:
    templateID = int(lbl_MRATemplate_ID.Content)      #dg_MRA_Templates.SelectedItem['TemplateID']
    caseTypeID = int(dg_MRATemplate_CaseTypes.SelectedItem['CaseTypeID'])
  except:
    MessageBox.Show("Error obtaining selected Matter Risk Assessment TemplateID or CaseType code", "Error: Unlink Case Type from Matter Risk Assessment Template...")
    return

  deleteSQL = """DELETE FROM Usr_MRAv2_CaseTypeDefaults 
                 WHERE TemplateID = {tID} AND CaseTypesCode = {ctID}""".format(tID=templateID, ctID=caseTypeID)
  try:
    _tikitResolver.Resolve("[SQL: {0}]".format(deleteSQL))
    refresh_MRA_CaseType_Defaults()
  except:
    MessageBox.Show("There was an error trying to unlink the selected Case Type from this Matter Risk Assessment Template, using SQL:\n" + str(deleteSQL), "Error: Unlink Case Type from Matter Risk Assessment Template...")
    return

  #MessageBox.Show("Remove Case Type to Template link button click", "Unlink Case Type from Matter Risk Assessment Template...")
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

## New MRA Edit Template tab code below ##

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


class GroupVM(NotifyBase):
  def __init__(self, group_name):
    NotifyBase.__init__(self)
    self._GroupName = group_name
    #self.Questions = ObservableCollection[QuestionVM]()  # forward ref ok in IronPython at runtime
    self.Questions = ObservableCollection[object]()

  @property
  def GroupName(self): return self._GroupName
  @GroupName.setter
  def GroupName(self, v):
      self._GroupName = v
      self._raise("GroupName")


class QuestionVM(NotifyBase):
  def __init__(self, question_id, text, question_display_order=0, parent_group=None):
    NotifyBase.__init__(self)
    self.QuestionID = question_id
    self.ParentGroup = parent_group
    self._QuestionText = text
    self._QuestionDisplayOrder = question_display_order
    #self.Answers = ObservableCollection[AnswerVM]()
    self.Answers = ObservableCollection[object]()

  @property
  def QuestionText(self): return self._QuestionText
  @QuestionText.setter
  def QuestionText(self, v):
      self._QuestionText = v
      self._raise("QuestionText")

  @property
  def QuestionDisplayOrder(self): return self._QuestionDisplayOrder
  @QuestionDisplayOrder.setter
  def QuestionDisplayOrder(self, v):
      self._QuestionDisplayOrder = v
      self._raise("QuestionDisplayOrder")


class AnswerVM(NotifyBase):
  def __init__(self, answer_id, text, score, email_comment='', answer_display_order=0, parent_question=None):
    NotifyBase.__init__(self)
    self.AnswerID = answer_id
    self.ParentQuestion = parent_question
    self._AnswerText = text
    self._Score = score
    self._EmailComment = email_comment
    self._AnswerDisplayOrder = answer_display_order

  @property
  def AnswerText(self): return self._AnswerText
  @AnswerText.setter
  def AnswerText(self, v):
      self._AnswerText = v
      self._raise("AnswerText")

  @property
  def Score(self): return self._Score
  @Score.setter
  def Score(self, v):
      self._Score = v
      self._raise("Score")

  @property
  def EmailComment(self): return self._EmailComment
  @EmailComment.setter
  def EmailComment(self, v):
      self._EmailComment = v
      self._raise("EmailComment")

  @property
  def AnswerDisplayOrder(self): return self._AnswerDisplayOrder
  @AnswerDisplayOrder.setter
  def AnswerDisplayOrder(self, v):
      self._AnswerDisplayOrder = v
      self._raise("AnswerDisplayOrder")


class TemplateEditorVM(NotifyBase):
  def __init__(self, template_id):
    NotifyBase.__init__(self)
    self.TemplateID = template_id
    #self.Groups = ObservableCollection[GroupVM]()
    self.Groups = ObservableCollection[object]()

    self._SelectedItem = None
    self._SelectedGroup = None
    self._SelectedQuestion = None
    self._SelectedAnswer = None

  @property
  def SelectedItem(self): return self._SelectedItem
  
  @SelectedItem.setter
  def SelectedItem(self, v):
      self._SelectedItem = v

      # Reset
      self._SelectedGroup = None
      self._SelectedQuestion = None
      self._SelectedAnswer = None

      if v is None:
        pass
      elif isinstance(v, GroupVM):
        self._SelectedGroup = v
        # optionally auto-select first question/answer?
        if v.Questions.Count > 0:
          self._SelectedQuestion = v.Questions[0]
          if self._SelectedQuestion.Answers.Count > 0:
            self._SelectedAnswer = self._SelectedQuestion.Answers[0]

      elif isinstance(v, QuestionVM):
        self._SelectedQuestion = v
        self._SelectedGroup = v.ParentGroup
        if v.Answers.Count > 0:
          self._SelectedAnswer = v.Answers[0]

      elif isinstance(v, AnswerVM):
        self._SelectedAnswer = v
        self._SelectedQuestion = v.ParentQuestion
        if self._SelectedQuestion is not None:
          self._SelectedGroup = self._SelectedQuestion.ParentGroup

      self._raise("SelectedItem")
      self._raise("SelectedGroup")
      self._raise("SelectedQuestion")
      self._raise("SelectedAnswer")

  @property
  def SelectedGroup(self): return self._SelectedGroup

  @property
  def SelectedQuestion(self): return self._SelectedQuestion
  
  @property
  def SelectedAnswer(self): return self._SelectedAnswer

## above mostly supplied from ChatGPT with minor modifications (adding additional fields recently added to XAML) ##

def tvTemplate_SelectedItemChanged(sender, e):
  # This function will handle when the selected item in the Template TreeView changes

  _tikitSender.DataContext.SelectedItem = e.NewValue

  #MessageBox.Show("Template TreeView selected item changed", "DEBUG - Template TreeView Selected Item Changed...")
  return


def load_template_structure_from_reader(vm, dr):
# This function will load the template structure from a data reader into the provided ViewModel (vm)

  group_map = {}  # group_name -> GroupVM
  q_map = {}      # (group_name, question_id) -> QuestionVM

  # Optional: if you want stable ordering
  group_names_in_order = []
  question_keys_in_order = []  # track first-seen order

  while dr.Read():
    group_name = "" if dr.IsDBNull(0) else dr.GetString(0)            # QuestionGroup
    q_order    = 0  if dr.IsDBNull(1) else int(dr.GetValue(1))        # QuestionOrder
    q_id       = 0  if dr.IsDBNull(2) else int(dr.GetValue(2))        # QuestionID
    q_text     = "" if dr.IsDBNull(3) else dr.GetString(3)            # QuestionText

    a_order    = 0  if dr.IsDBNull(4) else int(dr.GetValue(4))        # AnswerOrder
    a_id       = 0  if dr.IsDBNull(5) else int(dr.GetValue(5))        # AnswerID
    a_text     = "" if dr.IsDBNull(6) else dr.GetString(6)            # AnswerText
    email_c    = "" if dr.IsDBNull(7) else dr.GetString(7)            # EmailComment
    score      = 0  if dr.IsDBNull(8) else int(dr.GetValue(8))        # Score

    # Group
    g = group_map.get(group_name)
    if g is None:
      g = GroupVM(group_name)
      group_map[group_name] = g
      group_names_in_order.append(group_name)

    # Question
    q_key = (group_name, q_id)
    q = q_map.get(q_key)
    if q is None:
      q = QuestionVM(q_id, q_text, question_display_order=q_order, parent_group=g)
      q_map[q_key] = q
      g.Questions.Add(q)
      question_keys_in_order.append(q_key)
    else:
      # keep updated values (optional)
      q.QuestionText = q_text
      q.QuestionDisplayOrder = q_order

    # Answer (guard for LEFT JOIN nulls)
    if a_id != 0:
      a = AnswerVM(a_id, a_text, score, email_comment=email_c,
                    answer_display_order=a_order, parent_question=q)
      q.Answers.Add(a)

  # Now sort within each group/question if you want strict ordering
  # (ObservableCollection has no Sort, so rebuild in place)
  for group_name in group_names_in_order:
    g = group_map[group_name]

    # sort questions by display order
    qs = list(g.Questions)
    qs.sort(key=lambda qq: (qq.QuestionDisplayOrder, qq.QuestionText))
    g.Questions.Clear()
    for q in qs:
      # sort answers by answer display order
      ans = list(q.Answers)
      ans.sort(key=lambda aa: (aa.AnswerDisplayOrder, aa.AnswerText))
      q.Answers.Clear()
      for a in ans:
        q.Answers.Add(a)
      g.Questions.Add(q)

  # Load into vm
  vm.Groups.Clear()
  for group_name in sorted(group_map.keys()):  # or group_names_in_order for first-seen order
    vm.Groups.Add(group_map[group_name])

  # Set initial selection
  if vm.Groups.Count > 0:
    vm.SelectedItem = vm.Groups[0]


def EditMRA_loadTreeViewStructure(selectedTemplateID):

  if selectedTemplateID is None or selectedTemplateID == 0:
    return

  # form sql to get template structure
  sql = """SELECT MRAT.QuestionGroup, MRAT.QuestionOrder, MRAT.QuestionID, MRAQ.QuestionText, 
                  MRAT.AnswerOrder, MRAT.AnswerID, MRAA.AnswerText, MRAA.EmailComment, MRAT.Score
          FROM Usr_MRAv2_Templates MRAT
            JOIN Usr_MRAv2_TemplateDetails TD ON MRAT.TemplateID = TD.TemplateID
            LEFT OUTER JOIN Usr_MRAv2_Question MRAQ ON MRAT.QuestionID = MRAQ.QuestionID
            LEFT OUTER JOIN Usr_MRAv2_Answer MRAA ON MRAT.AnswerID = MRAA.AnswerID
          WHERE MRAT.TemplateID = {0}
          ORDER BY MRAT.QuestionGroup, MRAT.QuestionOrder, MRAT.AnswerOrder""".format(selectedTemplateID)

  vm = TemplateEditorVM(selectedTemplateID)
  _tikitSender.DataContext = vm

  _tikitDbAccess.Open(sql)
  dr = _tikitDbAccess._dr
  if dr is not None and dr.HasRows:
    load_template_structure_from_reader(vm, dr)
    dr.Close()
  _tikitDbAccess.Close()
  return


def next_question_order(group_vm):
  if group_vm is None or group_vm.Questions.Count == 0:
    return 1
  return max([q.QuestionDisplayOrder for q in group_vm.Questions]) + 1

def next_answer_order(question_vm):
  if question_vm is None or question_vm.Answers.Count == 0:
    return 1
  return max([a.AnswerDisplayOrder for a in question_vm.Answers]) + 1


# --- module-level temp negative ID generator for new items ---
def _next_temp_id():
  global _temp_id
  _temp_id -= 1
  return _temp_id

def _renumber_questions(group_vm):
  """Set QuestionDisplayOrder based on current list order (1..n)."""
  if group_vm is None: return
  i = 1
  for q in group_vm.Questions:
    q.QuestionDisplayOrder = i
    i += 1

def _renumber_answers(question_vm):
  """Set AnswerDisplayOrder based on current list order (1..n)."""
  if question_vm is None: return
  i = 1
  for a in question_vm.Answers:
    a.AnswerDisplayOrder = i
    i += 1

def _ensure_selected_context(vm):
  """
  If a user selects a Group but no SelectedQuestion/Answer set yet,
  populate sensible defaults. Your SelectedItem setter already does most of this,
  but this can help after programmatic inserts.
  """
  if vm is None: return
  if vm.SelectedGroup is None and vm.Groups.Count > 0:
    vm.SelectedItem = vm.Groups[0]


def btn_EditMRA_Question_Add_Click(sender, e):
  vm = _tikitSender.DataContext
  if vm is None:
    return

  # Determine target group
  g = vm.SelectedGroup
  if g is None:
    # Create a default group if nothing selected
    g = GroupVM("New Group")
    vm.Groups.Add(g)
    vm.SelectedItem = g

  # Create new question
  new_qid = _next_temp_id()
  q = QuestionVM(new_qid, "New question...", question_display_order=0, parent_group=g)
  g.Questions.Add(q)

  # Optional: create a starter answer so UI always has 3 levels populated
  new_aid = _next_temp_id()
  a = AnswerVM(new_aid, "New answer...", 0, email_comment="", answer_display_order=0, parent_question=q)
  q.Answers.Add(a)

  # Renumber display orders based on current list position
  _renumber_questions(g)
  _renumber_answers(q)

  # Select new question (or the answer if you prefer)
  vm.SelectedItem = q
  return

def btn_EditMRA_Answer_Add_Click(sender, e):
  
  vm = _tikitSender.DataContext
  if vm is None:
    return

  # Determine target question
  q = vm.SelectedQuestion
  if q is None and vm.SelectedAnswer is not None:
    q = vm.SelectedAnswer.ParentQuestion

  if q is None:
    # Nothing to attach answer to
    return

  new_aid = _next_temp_id()
  a = AnswerVM(new_aid, "New answer...", 0, email_comment="", answer_display_order=0, parent_question=q)
  q.Answers.Add(a)

  _renumber_answers(q)

  vm.SelectedItem = a
  return


######################################################################################################################


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
tb_MRA_NoneSelected = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MRA_NoneSelected')            # TextBlock shown when no MRA template selected
stk_ST_SelectedMRA = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'stk_ST_SelectedMRA')              # Edit Area StackPanel shown when MRA template selected
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

tb_ThisMRAid = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_ThisMRAid')
tb_ThisMRAname = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_ThisMRAname')

btn_BackToOverview_FromEditQs = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_BackToOverview_FromEditQs')
btn_BackToOverview_FromEditQs.Click += btn_BackToOverview_FromEditQs_Click

## clipboard area for copying questions/answers between MRAs  ##
grp_MRA_Clipboard = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'grp_MRA_Clipboard')
tb_MRA_CopiedWhat = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MRA_CopiedWhat')    # either: Question|One-Answer|All-Answers
tb_MRA_SourceTemplateID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MRA_SourceTemplateID')
tb_MRA_SourceQuestionID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MRA_SourceQuestionID')
tb_MRA_SourceAnswerID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MRA_SourceAnswerID')

## Toolbar Buttons for Editing Questions ##
btn_Questions_AddNew = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Questions_AddNew')
btn_Questions_AddNew.Click += btn_Questions_AddNew_Click
btn_Questions_Clipboard = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Questions_Clipboard')
btn_Questions_Clipboard.Click += btn_Questions_Clipboard_Click
QuestionsClipboard_Popup = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'QuestionsClipboard_Popup')
QuestionsClipboard_Popup.Closed += QuestionsClipboard_Popup_Closed
mi_Question_CopyToClipboard = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'mi_Question_CopyToClipboard')
mi_Question_CopyToClipboard.Click += mi_Question_CopyToClipboard_Click
mi_Question_PasteFromClipboard = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'mi_Question_PasteFromClipboard')
mi_Question_PasteFromClipboard.Click += mi_Question_PasteFromClipboard_Click

#btn_Questions_CopyToClipboard = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Questions_CopyToClipboard')
#btn_Questions_CopyToClipboard.Click += Duplicate_MRA_Question

btn_Question_MoveTop = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Question_MoveTop')
btn_Question_MoveTop.Click += btn_Question_MoveTop_Click
btn_Question_MoveUp = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Question_MoveUp')
btn_Question_MoveUp.Click += btn_Question_MoveUp_Click
btn_Question_MoveDown = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Question_MoveDown')
btn_Question_MoveDown.Click += btn_Question_MoveDown_Click
btn_Question_MoveBottom = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Question_MoveBottom')
btn_Question_MoveBottom.Click += btn_Question_MoveBottom_Click
btn_Question_DeleteSelected = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Question_DeleteSelected')
btn_Question_DeleteSelected.Click += btn_Question_DeleteSelected_Click

tb_MRA_NoQuestionsText = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MRA_NoQuestionsText')
dg_MRA_Questions = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_MRA_Questions')
dg_MRA_Questions.SelectionChanged += dg_MRA_Questions_SelectionChanged

## Editing Questions Area ##
tb_ESQ_QuestionID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_ESQ_QuestionID')
txt_ESQ_QuestionText = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'txt_ESQ_QuestionText')
txt_ESQ_QuestionGroup = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'txt_ESQ_QuestionGroup')
btn_ESQ_SaveQuestion = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_ESQ_SaveQuestion')
btn_ESQ_SaveQuestion.Click += btn_ESQ_SaveQuestion_Click

#########################################################################

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
#btn_MRAPreview_BackToOverview.Click += PreviewMRA_BackToOverview
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


## New MRA Template Editor ViewModel ##
#templateEditorVM = None
tvTemplate = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tvTemplate')      # tree view listing Question Group > Question > Answers
tvTemplate.SelectedItemChanged += tvTemplate_SelectedItemChanged
btn_EditMRA_Question_Add = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_EditMRA_Question_Add')
btn_EditMRA_Question_Add.Click += btn_EditMRA_Question_Add_Click
btn_EditMRA_Answer_Add = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_EditMRA_Answer_Add')
btn_EditMRA_Answer_Add.Click += btn_EditMRA_Answer_Add_Click


# Define Actions and on load events
myOnLoadEvent(_tikitSender, 'onLoad')

]]>
    </Loaded>
  </RiskPracticeV2>

</tfb>