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
from System import DateTime, Environment, String, Convert, DBNull, Action
from System.IO import Path, File, Directory
from System.Diagnostics import Process
from System.Globalization import DateTimeStyles
from System.Collections.Generic import Dictionary
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import Controls, Forms, LogicalTreeHelper, Clipboard 
from System.Windows import Data, UIElement, Visibility, Window
from System.Windows.Controls import Button, Canvas, GridView, GridViewColumn, ListView, Orientation, Validation, DataGridCellInfo
from System.Windows.Data import Binding, CollectionView, ListCollectionView, PropertyGroupDescription, CollectionViewSource
from System.Windows.Forms import SelectionMode, MessageBox, MessageBoxButtons, DialogResult, MessageBoxIcon
from System.Windows.Threading import DispatcherPriority
import re

## GLOBAL VARIABLES ##
preview_MRA = []    # To temp store table for previewing Matter Risk Assessment
_temp_id = -1

MRA_PREVIEW_ANSWERS_BY_QID = {}   # temp to store list of 'MRA_PREVIEW_ANSWER_ROW' dicts for current TemplateID being previewed
MRA_PREVIEW_QUESTIONS_LIST = []   # temp to store list of 'MRA_PREVIEW_QUESTION_ROW' dicts for current TemplateID being previewed
_preview_combo_syncing = False

## GLOBAL CONSTANTS##
UNSELECTED = -1

# ---------- Module-level clipboard ----------
MRA_CLIPBOARD = {
    "Mode": None,                 # "Q_ONLY" | "Q_AND_ALL_A" | "ALL_A" | "A_ONLY"
    "SourceTemplateID": None,
    "SourceQuestionID": None,
    "SourceAnswerID": None,
    "QuestionText": "",
    "Answers": []                 # list of dicts: {"AnswerText":..., "Score":..., "EmailComment":...}
}
# --------------------------------------------

# # # #   O N   L O A D   E V E N T   # # # #
def myOnLoadEvent(s, event):
  # populate drop-downs
  populate_MRATemplate_CaseTypeCombo()
  
  #MessageBox.Show("DEBUG - POPULATING LISTS")
  refresh_MRA_Templates()                      # Matter Risk Assessment (main overview / templates) 
  
  # hide 'Edit Questions' and 'Preview' tabs
  ti_MRA_Questions.Visibility = Visibility.Collapsed
  ti_MRA_Preview.Visibility = Visibility.Collapsed
  grp_MRA_Clipboard.Visibility = Visibility.Collapsed
  
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
                          'Q Count' = (SELECT COUNT(*) FROM (
                                          SELECT QuestionID FROM Usr_MRAv2_Templates MRAT 
                                          WHERE MRAT.TemplateID = TD.TemplateID GROUP BY MRAT.QuestionID) as tmpT)
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
  if dg_MRA_Templates.SelectedIndex == UNSELECTED:
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
  #nextTypeID = runSQL(codeToRun="SELECT ISNULL(MAX(TemplateID), 0) + 1 FROM Usr_MRAv2_TemplateDetails", returnType='Int')
  insertSQL = """INSERT INTO Usr_MRAv2_TemplateDetails (Name, DaysUntil_IncompleteLock, ScoreMediumTrigger, ScoreHighTrigger, TemplateID)
                 OUTPUT INSERTED.TemplateID
                 SELECT 'NMRA - new', 29, 0, 0, ISNULL(MAX(TemplateID), 0) + 1 FROM Usr_MRAv2_TemplateDetails"""
  
  try:
    nextTypeID = _tikitResolver.Resolve("[SQL: {0}]".format(insertSQL))
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
    currentSelectedID = dg_MRA_Templates.SelectedItem['TemplateID'] if dg_MRA_Templates.SelectedIndex != UNSELECTED else None

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
  vm = _tikitSender.DataContext
  # note: vm.Debug to populate our on-screen debug text box, and 'log_line' to write to users daily log file
  if vm is None:
    vm.Debug("ERROR: Unable to load MRA TemplateID {0} into TreeView structure for editing".format(tb_ThisMRAid.Text))  
    log_line("ERROR: Unable to load MRA TemplateID {0} into TreeView structure for editing".format(tb_ThisMRAid.Text))
  else:
    vm.Debug("Loaded MRA TemplateID {0} into TreeView structure for editing".format(vm.TemplateID))
    log_line("Loaded MRA TemplateID {0} into TreeView structure for editing".format(vm.TemplateID))
    # also output a 'before' dump of view model to file for audit purposes
    dump_vm_to_file(vm, prefix="AUTO_BEFORE")

  # show 'Questions' tab and hide 'Overview' tab
  ti_MRA_Overview.Visibility = Visibility.Collapsed
  ti_MRA_Questions.Visibility = Visibility.Visible
  ti_MRA_Questions.IsSelected = True
  #MessageBox.Show("EditSelected_Click", "DEBUG - TESTING")
  return

# # # #  END OF:  Matter Risk Assessment Templates   # # # #


def btn_Clipboard_Copy_Click(s, event):
  # This function will open the 'Questions Clipboard' window for copying/pasting questions between templates
  
  Clipboard_Popup_Copy.IsOpen = btn_Clipboard_Copy.IsChecked
  #MessageBox.Show("Questions Clipboard button click", "Open Questions Clipboard...")
  return

def Clipboard_Popup_Copy_Closed(s, event):
  # This function will uncheck the 'Questions Clipboard' button when the popup is closed
  
  btn_Clipboard_Copy.IsChecked = False
  Clipboard_Popup_Copy.IsOpen = False
  return

def _answer_to_dict(a):
  return {
      "AnswerText": a.AnswerText or "",
      "Score": int(a.Score) if a.Score is not None and str(a.Score) != "" else 0,
      "EmailComment": a.EmailComment or ""
    }

def _set_clipboard_indicators(mode, tid, qid, aid):
  tb_copied = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MRA_CopiedWhat')
  tb_tid    = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MRA_SourceTemplateID')
  tb_qid    = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MRA_SourceQuestionID')
  tb_aid    = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MRA_SourceAnswerID')

  if tb_copied is not None: tb_copied.Text = str(mode or "")
  if tb_tid is not None:    tb_tid.Text = str(tid if tid is not None else -1)
  if tb_qid is not None:    tb_qid.Text = str(qid if qid is not None else -1)
  if tb_aid is not None:    tb_aid.Text = str(aid if aid is not None else -1)
  # show area
  grp_MRA_Clipboard.Visibility = Visibility.Visible
  return

def _add_answers_to_question(vm, q, answer_dicts):
  if q is None:
    return

  # local temp id generator (same approach you used earlier)
  # assumes you have _next_temp_id() implemented
  for ad in answer_dicts:
      aid = _next_temp_id()
      a = AnswerVM(
          aid,
          ad.get("AnswerText", ""),
          ad.get("Score", 0),
          email_comment=ad.get("EmailComment", ""),
          answer_display_order=0,
          parent_question=q
      )
      q.Answers.Add(a)

  _renumber_answers(q)
  return

def btn_CopyToClipboard_QuestionOnly_Click(s, event):
  # This function will copy only the question text to the clipboard

  vm = _tikitSender.DataContext
  q = vm.SelectedQuestion
  if q is None:
    return

  MRA_CLIPBOARD["Mode"] = "Q_ONLY"
  MRA_CLIPBOARD["SourceTemplateID"] = vm.TemplateID
  MRA_CLIPBOARD["SourceQuestionID"] = q.QuestionID
  MRA_CLIPBOARD["SourceAnswerID"] = None
  MRA_CLIPBOARD["QuestionText"] = q.QuestionText or ""
  MRA_CLIPBOARD["Answers"] = []

  _set_clipboard_indicators("Question only", vm.TemplateID, q.QuestionID, None)

  Clipboard_Popup_Copy_Closed(s, event)
  #MessageBox.Show("Copied Question Only to Clipboard", "Copied Question Only to Clipboard...")
  # logging instead of messagebox for this one as it's more likely to be used when copying multiple questions, and we don't want to bombard user with message boxes in that scenario
  vm.Debug("Copied question (ID={0}) to clipboard".format(q.QuestionID))
  log_line("Copied question (ID={0}) to clipboard".format(q.QuestionID))
  return

def btn_CopyToClipboard_QuestionAndAnswers_Click(s, event):
  # This function will copy the question and all possible answers to the clipboard
  
  vm = _tikitSender.DataContext
  q = vm.SelectedQuestion
  if q is None:
    return

  MRA_CLIPBOARD["Mode"] = "Q_AND_ALL_A"
  MRA_CLIPBOARD["SourceTemplateID"] = vm.TemplateID
  MRA_CLIPBOARD["SourceQuestionID"] = q.QuestionID
  MRA_CLIPBOARD["SourceAnswerID"] = None
  MRA_CLIPBOARD["QuestionText"] = q.QuestionText or ""
  MRA_CLIPBOARD["Answers"] = [_answer_to_dict(a) for a in q.Answers]

  _set_clipboard_indicators("Question + all answers", vm.TemplateID, q.QuestionID, None)
  
  Clipboard_Popup_Copy_Closed(s, event)
  #MessageBox.Show("Copied Question and Answers to Clipboard", "Copied Question and Answers to Clipboard...")
  vm.Debug("Copied question (ID={0}) and all answers to clipboard".format(q.QuestionID))
  log_line("Copied question (ID={0}) and all answers to clipboard".format(q.QuestionID))
  return

def btn_CopyToClipboard_AnswersAll_Click(s, event):
  # This function will copy all possible answers to the clipboard
  
  vm = _tikitSender.DataContext
  q = vm.SelectedQuestion
  if q is None:
    return

  MRA_CLIPBOARD["Mode"] = "ALL_A"
  MRA_CLIPBOARD["SourceTemplateID"] = vm.TemplateID
  MRA_CLIPBOARD["SourceQuestionID"] = q.QuestionID
  MRA_CLIPBOARD["SourceAnswerID"] = None
  MRA_CLIPBOARD["QuestionText"] = ""
  MRA_CLIPBOARD["Answers"] = [_answer_to_dict(a) for a in q.Answers]

  _set_clipboard_indicators("All answers", vm.TemplateID, q.QuestionID, None)

  Clipboard_Popup_Copy_Closed(s, event)
  #MessageBox.Show("Copied All Answers to Clipboard", "Copied All Answers to Clipboard...")
  vm.Debug("Copied all answers for question (ID={0}) to clipboard".format(q.QuestionID))
  log_line("Copied all answers for question (ID={0}) to clipboard".format(q.QuestionID))
  return

def btn_CopyToClipboard_AnswerOnly_Click(s, event):
  # This function will copy only the selected answer to the clipboard

  vm = _tikitSender.DataContext
  a = vm.SelectedAnswer
  q = vm.SelectedQuestion
  if a is None or q is None:
    return

  MRA_CLIPBOARD["Mode"] = "A_ONLY"
  MRA_CLIPBOARD["SourceTemplateID"] = vm.TemplateID
  MRA_CLIPBOARD["SourceQuestionID"] = q.QuestionID
  MRA_CLIPBOARD["SourceAnswerID"] = a.AnswerID
  MRA_CLIPBOARD["QuestionText"] = ""
  MRA_CLIPBOARD["Answers"] = [_answer_to_dict(a)]

  _set_clipboard_indicators("Single answer", vm.TemplateID, q.QuestionID, a.AnswerID)  
  
  Clipboard_Popup_Copy_Closed(s, event)
  #MessageBox.Show("Copied Answer Only to Clipboard", "Copied Answer Only to Clipboard...")
  vm.Debug("Copied answer (ID={0}) to clipboard".format(a.AnswerID))
  log_line("Copied answer (ID={0}) to clipboard".format(a.AnswerID))
  return


def btn_CopyToClipboard_Clear_Click(s, event):
  # This function will clear the clipboard and hide the indicators

  MRA_CLIPBOARD["Mode"] = None
  MRA_CLIPBOARD["SourceTemplateID"] = None
  MRA_CLIPBOARD["SourceQuestionID"] = None
  MRA_CLIPBOARD["SourceAnswerID"] = None
  MRA_CLIPBOARD["QuestionText"] = ""
  MRA_CLIPBOARD["Answers"] = []

  vm = _tikitSender.DataContext

  # hide area
  _set_clipboard_indicators("", None, None, None)
  grp_MRA_Clipboard.Visibility = Visibility.Collapsed
  #MessageBox.Show("Clipboard has been emptied", "Cleared Clipboard...")
  vm.Debug("Clipboard cleared")
  log_line("Clipboard cleared")
  return

def btn_Clipboard_Paste_Click(s, event):
  # This function will paste the contents of the clipboard into the selected question/answer area
  vm = _tikitSender.DataContext
  clipboard_paste(vm)
  return

def clipboard_paste(vm):
  mode = MRA_CLIPBOARD.get("Mode")
  if not mode:
    MessageBox.Show("Clipboard is empty.", "Paste", MessageBoxButtons.OK, MessageBoxIcon.Information)
    return

  # Ensure group exists for new question creation
  g = vm.SelectedGroup
  if g is None and vm.Groups.Count > 0:
    g = vm.Groups[0]

  # Determine if we should paste into selected or create new
  q_selected = vm.SelectedQuestion

  contains_qtext = mode in ("Q_ONLY", "Q_AND_ALL_A")
  contains_answers = mode in ("Q_AND_ALL_A", "ALL_A", "A_ONLY")

  target_q = q_selected

  # If question text is included, allow overwrite or create new
  if contains_qtext:
    if target_q is not None:
        res = MessageBox.Show(
            "Paste question text into the selected question?\n\nYes = overwrite selected\nNo = create new question",
            "Paste Question",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question
        )
        if res == DialogResult.No:
          target_q = None

    if target_q is None:
        # create a new question in selected group
        if g is None:
            g = GroupVM("New Group")
            vm.Groups.Add(g)

        qid = _next_temp_id()
        target_q = QuestionVM(qid, "New question...", question_display_order=0, parent_group=g)
        g.Questions.Add(target_q)
        _renumber_questions(g)

    # Apply question text
    target_q.QuestionText = MRA_CLIPBOARD.get("QuestionText", "")

  # If answers are included, decide where they go
  if contains_answers:
    if target_q is None:
        # no selected question and we didn't paste/create one above
        if g is None:
            g = GroupVM("New Group")
            vm.Groups.Add(g)

        qid = _next_temp_id()
        target_q = QuestionVM(qid, "New question...", question_display_order=0, parent_group=g)
        g.Questions.Add(target_q)
        _renumber_questions(g)

    # Optional: ask append vs replace
    if target_q.Answers.Count > 0:
        res = MessageBox.Show(
            "Add copied answers to the selected question?\n\nYes = append\nNo = replace existing answers",
            "Paste Answers",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question
        )
        if res == DialogResult.No:
            target_q.Answers.Clear()

    _add_answers_to_question(vm, target_q, MRA_CLIPBOARD.get("Answers", []))

  # Select something sensible after paste
  vm.SelectedItem = target_q
  # output paste operation to debug and log
  vm.Debug("Pasted {0} from TemplateID={1}, QuestionID={2}, AnswerID={3}".format(
      mode,
      MRA_CLIPBOARD.get("SourceTemplateID"),
      MRA_CLIPBOARD.get("SourceQuestionID"),
      MRA_CLIPBOARD.get("SourceAnswerID")
  ))
  log_line("Pasted {0} from TemplateID={1}, QuestionID={2}, AnswerID={3}".format(
      mode,
      MRA_CLIPBOARD.get("SourceTemplateID"),
      MRA_CLIPBOARD.get("SourceQuestionID"),
      MRA_CLIPBOARD.get("SourceAnswerID")
  ))
  return

 
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
## We have a few classes to represent the ViewModel for the Template Editor ##
# These classes implement INotifyPropertyChanged to support data binding in WPF
# We also use ObservableCollection for collections that can change dynamically
# GroupVM - represents a group of questions
# QuestionVM - represents a question with its answers
# AnswerVM - represents an answer to a question
# TemplateEditorVM - represents the overall template editor view model

class NotifyBase(INotifyPropertyChanged):
  # This creates add_PropertyChanged/remove_PropertyChanged automatically
  def __init__(self):
    # store delegates that WPF adds via add_PropertyChanged
    self._pc_handlers = []
    self._log = None  # function(str)

  def SetLogger(self, log_func):
    self._log = log_func

  def _log_change(self, msg):
    if callable(self._log):
      self._log(msg)


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
    self.NodeKind = "Group"
    self._GroupName = group_name
    self.Questions = ObservableCollection[object]()  # ideally 'QuestionVM' instead of 'object', but forward ref isn't liked in IronPython in this context, so using 'object' to avoid IDE errors (will still be 'QuestionVM' at runtime)

  @property
  def GroupName(self): return self._GroupName
  @GroupName.setter
  def GroupName(self, v):
      newV = "" if v is None else str(v)
      oldV = "" if self._GroupName is None else str(self._GroupName)
      if newV == oldV:
        return
      self._GroupName = newV
      self._raise("GroupName")
      self._log_change("GroupName changed: '{0}' -> '{1}'".format(oldV, newV))


class QuestionVM(NotifyBase):
  def __init__(self, question_id, text, question_display_order=0, parent_group=None):
    NotifyBase.__init__(self)
    self.NodeKind = "Question"
    self.QuestionID = question_id
    self.ParentGroup = parent_group
    self._QuestionText = text or ""
    self._QuestionDisplayOrder = question_display_order
    self.Answers = ObservableCollection[object]()   # ideally 'AnswerVM' instead of 'object', but forward ref isn't liked in IronPython in this context, so using 'object' to avoid IDE errors (will still be 'QuestionVM' at runtime)

  @property
  def QuestionText(self): return self._QuestionText
  @QuestionText.setter
  def QuestionText(self, v):
      newV = "" if v is None else str(v)
      oldV = "" if self._QuestionText is None else str(self._QuestionText)
      if newV == oldV:
        return
      self._QuestionText = newV
      self._raise("QuestionText")
      self._log_change("QuestionText changed - QID={0}: '{1}' -> '{2}'".format(self.QuestionID, oldV, newV))

  @property
  def QuestionDisplayOrder(self): return self._QuestionDisplayOrder
  @QuestionDisplayOrder.setter
  def QuestionDisplayOrder(self, v):
      newV = 0 if v is None else int(v)
      oldV = 0 if self._QuestionDisplayOrder is None else int(self._QuestionDisplayOrder)
      if newV == oldV:
        return
      self._QuestionDisplayOrder = newV
      self._raise("QuestionDisplayOrder")
      self._log_change("QuestionDisplayOrder changed - QID={0}: {1} -> {2}".format(self.QuestionID, oldV, newV))


class AnswerVM(NotifyBase):
  def __init__(self, answer_id, text, score, email_comment='', answer_display_order=0, parent_question=None):
    NotifyBase.__init__(self)
    self.NodeKind = "Answer"
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
      newV = "" if v is None else str(v)
      oldV = "" if self._AnswerText is None else str(self._AnswerText)
      if newV == oldV:
        return
      self._AnswerText = newV
      self._raise("AnswerText")
      self._log_change("AnswerText changed - AID={0}: '{1}' -> '{2}'".format(self.AnswerID, oldV, newV))

  @property
  def Score(self): return self._Score
  @Score.setter
  def Score(self, v):
      newV = 0 if v is None else int(v)
      oldV = 0 if self._Score is None else int(self._Score)
      if newV == oldV:
        return
      self._Score = newV
      self._raise("Score")
      self._log_change("Score changed - AID={0}: {1} -> {2}".format(self.AnswerID, oldV, newV))

  @property
  def EmailComment(self): return self._EmailComment
  @EmailComment.setter
  def EmailComment(self, v):
      newV = "" if v is None else str(v)
      oldV = "" if self._EmailComment is None else str(self._EmailComment)
      if newV == oldV:
        return
      self._EmailComment = newV
      self._raise("EmailComment")
      self._log_change("EmailComment changed - AID={0}: '{1}' -> '{2}'".format(self.AnswerID, oldV, newV))

  @property
  def AnswerDisplayOrder(self): return self._AnswerDisplayOrder
  @AnswerDisplayOrder.setter
  def AnswerDisplayOrder(self, v):
      newV = 0 if v is None else int(v)
      oldV = 0 if self._AnswerDisplayOrder is None else int(self._AnswerDisplayOrder)
      if newV == oldV:
        return
      self._AnswerDisplayOrder = newV
      self._raise("AnswerDisplayOrder")
      self._log_change("AnswerDisplayOrder changed - AID={0}: {1} -> {2}".format(self.AnswerID, oldV, newV))


class TemplateEditorVM(NotifyBase):
  def __init__(self, template_id):
    NotifyBase.__init__(self)
    self.TemplateID = template_id
    #self.Groups = ObservableCollection[GroupVM]()
    self.Groups = ObservableCollection[object]()

    self.DebugLines = ObservableCollection[String]() # for debugging output (list of strings)
    self._debug_max = 400

    self._SelectedItem = None
    self._SelectedGroup = None
    self._SelectedQuestion = None
    self._SelectedAnswer = None

  def Debug(self, msg):
    # timestamp optional; keep it simple
    self.DebugLines.Add(String.Format("{0}", msg))
    if self.DebugLines.Count > self._debug_max:
      self.DebugLines.RemoveAt(0)
    self._raise("DebugLines")

  @property
  def SelectedItem(self): return self._SelectedItem
  
  @SelectedItem.setter
  def SelectedItem(self, v):
      self._SelectedItem = v

      # Reset
      self._SelectedGroup = None
      self._SelectedQuestion = None
      self._SelectedAnswer = None

      # Helpers: safely detect by attribute presence (works with proxy wrappers)
      def has(obj, name):
          try:
              getattr(obj, name)
              return True
          except:
              return False

      kind = None
      try:
          kind = v.NodeKind
      except:
          kind = None

      if v is None:
          pass

      elif kind == "Group":
          self._SelectedGroup = v
          try:
              if v.Questions is not None and v.Questions.Count > 0:
                  self._SelectedQuestion = v.Questions[0]
                  if self._SelectedQuestion.Answers is not None and self._SelectedQuestion.Answers.Count > 0:
                      self._SelectedAnswer = self._SelectedQuestion.Answers[0]
          except:
              pass

      elif kind == "Question":
          self._SelectedQuestion = v
          try:
              self._SelectedGroup = v.ParentGroup
          except:
              self._SelectedGroup = None
          try:
              if v.Answers is not None and v.Answers.Count > 0:
                  self._SelectedAnswer = v.Answers[0]
          except:
              pass

      elif kind == "Answer":
          self._SelectedAnswer = v
          try:
              self._SelectedQuestion = v.ParentQuestion
          except:
              self._SelectedQuestion = None
          try:
              if self._SelectedQuestion is not None:
                  self._SelectedGroup = self._SelectedQuestion.ParentGroup
          except:
              self._SelectedGroup = None

      else:
          # Unknown projection object: fall back to attribute checks (optional but useful)
          # This makes it resilient if something ever comes through without NodeKind.

          # Answer-like: has AnswerID and ParentQuestion
          if has(v, "AnswerID") and has(v, "ParentQuestion"):
              self._SelectedAnswer = v
              try:
                  self._SelectedQuestion = v.ParentQuestion
              except:
                  self._SelectedQuestion = None
              if self._SelectedQuestion is not None:
                  try:
                      self._SelectedGroup = self._SelectedQuestion.ParentGroup
                  except:
                      self._SelectedGroup = None

          # Question-like: has QuestionID and Answers and ParentGroup
          elif has(v, "QuestionID") and has(v, "Answers") and has(v, "ParentGroup"):
              self._SelectedQuestion = v
              try:
                  self._SelectedGroup = v.ParentGroup
              except:
                  self._SelectedGroup = None
              try:
                  if v.Answers is not None and v.Answers.Count > 0:
                      self._SelectedAnswer = v.Answers[0]
              except:
                  pass

          # Group-like: has GroupName and Questions
          elif has(v, "GroupName") and has(v, "Questions"):
              self._SelectedGroup = v
              try:
                  if v.Questions is not None and v.Questions.Count > 0:
                      self._SelectedQuestion = v.Questions[0]
                      if self._SelectedQuestion.Answers is not None and self._SelectedQuestion.Answers.Count > 0:
                          self._SelectedAnswer = self._SelectedQuestion.Answers[0]
              except:
                  pass


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

def attach_logger(vm):
  # Create one logger function for this VM/session
  def vm_logger(msg):
    # update daily log file
    log_line(msg, vm.TemplateID)
    # 2) also push to on-screen debug (so we can confirm setters fire)
    try:
      vm.Debug(str(msg))
    except:
      pass

  # Store it on the vm so you can reuse it for newly-created objects
  vm.SetLogger(vm_logger)

  # Also useful: keep a reference explicitly (clearer than accessing vm._log)
  vm.Logger = vm_logger

  # Push logger into existing nodes
  for g in vm.Groups:
    g.SetLogger(vm_logger)
    for q in g.Questions:
      q.SetLogger(vm_logger)
      for a in q.Answers:
        a.SetLogger(vm_logger)

  return

## above mostly supplied from ChatGPT with minor modifications (adding additional fields recently added to XAML) ##

def tvTemplate_SelectedItemChanged(sender, e):
  # This function will handle when the selected item in the Template TreeView changes

  vm = _tikitSender.DataContext
  if vm is None:
      return
  vm.SelectedItem = e.NewValue

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

  attach_logger(vm)
  return


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

  def vm_logger(msg):
    log_line(msg, selectedTemplateID)
  vm.SetLogger(vm_logger)
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
  # Set QuestionDisplayOrder based on current list order (1..n).
  if group_vm is None: return
  i = 1
  for q in group_vm.Questions:
    q.QuestionDisplayOrder = i
    i += 1
  return

def _renumber_answers(question_vm):
  # Set AnswerDisplayOrder based on current list order (1..n).
  if question_vm is None: return
  i = 1
  for a in question_vm.Answers:
    a.AnswerDisplayOrder = i
    i += 1
  return

def _move_item_to_index(collection, item, new_index):
  # Moves item to new_index within an ObservableCollection.
  if collection is None or item is None:
    return False

  old_index = collection.IndexOf(item)
  if old_index < 0:
    return False

  # clamp
  if new_index < 0:
    new_index = 0
  if new_index > collection.Count - 1:
    new_index = collection.Count - 1

  if new_index == old_index:
    return False

  # ObservableCollection has Move(oldIndex,newIndex) in WPF
  collection.Move(old_index, new_index)
  return True

def _ensure_selected_context(vm):
  #  If a user selects a Group but no SelectedQuestion/Answer set yet,
  # populate sensible defaults. Your SelectedItem setter already does most of this,
  # but this can help after programmatic inserts.

  if vm is None: return
  if vm.SelectedGroup is None and vm.Groups.Count > 0:
    vm.SelectedItem = vm.Groups[0]


def btn_EditMRA_Group_Add_Click(sender, e):
  vm = _tikitSender.DataContext
  if vm is None:
    return

  new_group = GroupVM("New Group")
  # inherit logger
  if callable(getattr(vm, 'Logger', None)):
    new_group.SetLogger(vm.Logger)

  vm.Groups.Add(new_group)
  vm.SelectedItem = new_group

  vm.Debug("Added new group '{0}' with default question/answer".format(new_group.GroupName))
  log_line("Added new group '{0}' with default question/answer".format(new_group.GroupName))
  # we ought to add dummy question/answer too so UI has 3 levels populated
  EditMRA_Question_AddNew(vm)
  return


def btn_EditMRA_Question_Add_Click(sender, e):
  vm = _tikitSender.DataContext
  if vm is None:
    return

  EditMRA_Question_AddNew(vm)
  return


def EditMRA_Question_AddNew(vm):
  # Determine target group
  g = vm.SelectedGroup
  if g is None:
    # Create a default group if nothing selected
    g = GroupVM("New Group")
    if callable(getattr(vm, '_log', None)):
      g.SetLogger(vm._log)

    vm.Groups.Add(g)
    vm.SelectedItem = g

  # Create new question
  new_qid = _next_temp_id()
  q = QuestionVM(new_qid, "New question...", question_display_order=0, parent_group=g)
  if callable(getattr(vm, '_log', None)):
    q.SetLogger(vm._log)
  
  g.Questions.Add(q)

  # Optional: create a starter answer so UI always has 3 levels populated
  new_aid = _next_temp_id()
  a = AnswerVM(new_aid, "New answer...", 0, email_comment="", answer_display_order=0, parent_question=q)
  if callable(getattr(vm, '_log', None)):
    a.SetLogger(vm._log)

  q.Answers.Add(a)

  # Renumber display orders based on current list position
  _renumber_questions(g)
  _renumber_answers(q)

  # Select new question (or the answer if you prefer)
  vm.SelectedItem = q
  vm.Debug("Added new question '{0}' to group '{1}'".format(q.QuestionText, g.GroupName))
  log_line("Added new question '{0}' to group '{1}'".format(q.QuestionText, g.GroupName))
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
  if callable(getattr(vm, '_log', None)):
    a.SetLogger(vm._log)

  q.Answers.Add(a)

  _renumber_answers(q)

  vm.SelectedItem = a
  vm.Debug("Added new answer '{0}' to question '{1}'".format(a.AnswerText, q.QuestionText))
  log_line("Added new answer '{0}' to question '{1}'".format(a.AnswerText, q.QuestionText))
  return


def btn_EditMRA_Group_MoveTop_Click(sender, e):
  vm = _tikitSender.DataContext
  g = vm.SelectedGroup
  if g is None:
    return

  if _move_item_to_index(vm.Groups, g, 0):
    vm.SelectedItem = g  # keep selection stable
  vm.Debug("Moved group '{0}' to top".format(g.GroupName))
  log_line("Moved group '{0}' to top".format(g.GroupName))
  return

def btn_EditMRA_Group_MoveUp_Click(sender, e):
  vm = _tikitSender.DataContext
  g = vm.SelectedGroup
  if g is None:
    return

  idx = vm.Groups.IndexOf(g)
  if idx <= 0:
    return

  if _move_item_to_index(vm.Groups, g, idx - 1):
    vm.SelectedItem = g
  vm.Debug("Moved group '{0}' up".format(g.GroupName))
  log_line("Moved group '{0}' up".format(g.GroupName))
  return

def btn_EditMRA_Group_MoveDown_Click(sender, e):
  vm = _tikitSender.DataContext
  g = vm.SelectedGroup
  if g is None:
    return

  idx = vm.Groups.IndexOf(g)
  if idx < 0 or idx >= vm.Groups.Count - 1:
    return

  if _move_item_to_index(vm.Groups, g, idx + 1):
    vm.SelectedItem = g
  vm.Debug("Moved group '{0}' down".format(g.GroupName))
  log_line("Moved group '{0}' down".format(g.GroupName))
  return

def btn_EditMRA_Group_MoveBottom_Click(sender, e):
  vm = _tikitSender.DataContext
  g = vm.SelectedGroup
  if g is None:
    return

  if _move_item_to_index(vm.Groups, g, vm.Groups.Count - 1):
    vm.SelectedItem = g
  vm.Debug("Moved group '{0}' to bottom".format(g.GroupName))
  log_line("Moved group '{0}' to bottom".format(g.GroupName))
  return

def btn_EditMRA_Group_DeleteSelected_Click(sender, e):
  vm = _tikitSender.DataContext
  g = vm.SelectedGroup
  if g is None:
    return

  # Confirm deletion
  res = MessageBox.Show("Are you sure you want to delete the selected group and all its questions and answers?", "Confirm Delete Group", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
  if res != DialogResult.Yes:
    return

  vm.Groups.Remove(g)
  _ensure_selected_context(vm)

  vm.Debug("Deleted group '{0}'".format(g.GroupName))
  log_line("Deleted group '{0}'".format(g.GroupName))
  return


def btn_EditMRA_Question_MoveTop_Click(sender, e):
  vm = _tikitSender.DataContext
  q = vm.SelectedQuestion
  g = vm.SelectedGroup
  if g is None or q is None:
    return

  if _move_item_to_index(g.Questions, q, 0):
    _renumber_questions(g)
    vm.SelectedItem = q  # keep selection stable
  
  vm.Debug("Moved question '{0}' to top of group '{1}'".format(q.QuestionText, g.GroupName))
  log_line("Moved question '{0}' to top of group '{1}'".format(q.QuestionText, g.GroupName))  
  return

def btn_EditMRA_Question_MoveUp_Click(sender, e):
  vm = _tikitSender.DataContext
  q = vm.SelectedQuestion
  g = vm.SelectedGroup
  if g is None or q is None:
    return

  idx = g.Questions.IndexOf(q)
  if idx <= 0:
    return

  if _move_item_to_index(g.Questions, q, idx - 1):
    _renumber_questions(g)
    vm.SelectedItem = q
  vm.Debug("Moved question '{0}' up in group '{1}'".format(q.QuestionText, g.GroupName))
  log_line("Moved question '{0}' up in group '{1}'".format(q.QuestionText, g.GroupName))
  return

def btn_EditMRA_Question_MoveDown_Click(sender, e):
  vm = _tikitSender.DataContext
  q = vm.SelectedQuestion
  g = vm.SelectedGroup
  if g is None or q is None:
    return

  idx = g.Questions.IndexOf(q)
  if idx < 0 or idx >= g.Questions.Count - 1:
    return

  if _move_item_to_index(g.Questions, q, idx + 1):
    _renumber_questions(g)
    vm.SelectedItem = q
  vm.Debug("Moved question '{0}' down in group '{1}'".format(q.QuestionText, g.GroupName))
  log_line("Moved question '{0}' down in group '{1}'".format(q.QuestionText, g.GroupName))
  return

def btn_EditMRA_Question_MoveBottom_Click(sender, e):
  vm = _tikitSender.DataContext
  q = vm.SelectedQuestion
  g = vm.SelectedGroup
  if g is None or q is None:
    return

  if _move_item_to_index(g.Questions, q, g.Questions.Count - 1):
    _renumber_questions(g)
    vm.SelectedItem = q
  vm.Debug("Moved question '{0}' to bottom of group '{1}'".format(q.QuestionText, g.GroupName))
  log_line("Moved question '{0}' to bottom of group '{1}'".format(q.QuestionText, g.GroupName))
  return


def btn_EditMRA_Question_DeleteSelected_Click(sender, e):
  vm = _tikitSender.DataContext
  q = vm.SelectedQuestion
  g = vm.SelectedGroup
  if g is None or q is None:
    return

  # Confirm deletion
  res = MessageBox.Show("Are you sure you want to delete the selected question and all its answers?", "Confirm Delete Question", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
  if res != DialogResult.Yes:
    return

  g.Questions.Remove(q)
  _renumber_questions(g)
  _ensure_selected_context(vm)

  vm.Debug("Deleted question '{0}' from group '{1}'".format(q.QuestionText, g.GroupName))
  log_line("Deleted question '{0}' from group '{1}'".format(q.QuestionText, g.GroupName))
  return


def btn_EditMRA_Answer_MoveTop_Click(sender, e):
  vm = _tikitSender.DataContext
  a = vm.SelectedAnswer
  q = vm.SelectedQuestion
  if q is None or a is None:
    return

  if _move_item_to_index(q.Answers, a, 0):
    _renumber_answers(q)
    vm.SelectedItem = a

  vm.Debug("Moved answer '{0}' to top of question '{1}'".format(a.AnswerText, q.QuestionText))
  log_line("Moved answer '{0}' to top of question '{1}'".format(a.AnswerText, q.QuestionText))
  return

def btn_EditMRA_Answer_MoveUp_Click(sender, e):
  vm = _tikitSender.DataContext
  a = vm.SelectedAnswer
  q = vm.SelectedQuestion
  if q is None or a is None:
    return

  idx = q.Answers.IndexOf(a)
  if idx <= 0:
    return

  if _move_item_to_index(q.Answers, a, idx - 1):
    _renumber_answers(q)
    vm.SelectedItem = a

  vm.Debug("Moved answer '{0}' up in question '{1}'".format(a.AnswerText, q.QuestionText))
  log_line("Moved answer '{0}' up in question '{1}'".format(a.AnswerText, q.QuestionText))
  return


def btn_EditMRA_Answer_MoveDown_Click(sender, e):
  vm = _tikitSender.DataContext
  a = vm.SelectedAnswer
  q = vm.SelectedQuestion
  if q is None or a is None:
    return

  idx = q.Answers.IndexOf(a)
  if idx < 0 or idx >= q.Answers.Count - 1:
    return

  if _move_item_to_index(q.Answers, a, idx + 1):
    _renumber_answers(q)
    vm.SelectedItem = a

  vm.Debug("Moved answer '{0}' down in question '{1}'".format(a.AnswerText, q.QuestionText))
  log_line("Moved answer '{0}' down in question '{1}'".format(a.AnswerText, q.QuestionText))
  return

def btn_EditMRA_Answer_MoveBottom_Click(sender, e):
  vm = _tikitSender.DataContext
  a = vm.SelectedAnswer
  q = vm.SelectedQuestion
  if q is None or a is None:
    return

  if _move_item_to_index(q.Answers, a, q.Answers.Count - 1):
    _renumber_answers(q)
    vm.SelectedItem = a

  vm.Debug("Moved answer '{0}' to bottom of question '{1}'".format(a.AnswerText, q.QuestionText))
  log_line("Moved answer '{0}' to bottom of question '{1}'".format(a.AnswerText, q.QuestionText))
  return


def btn_EditMRA_Answer_DeleteSelected_Click(sender, e):
  vm = _tikitSender.DataContext
  a = vm.SelectedAnswer
  q = vm.SelectedQuestion
  if q is None or a is None:
    return

  # Confirm deletion
  res = MessageBox.Show("Are you sure you want to delete the selected answer?", "Confirm Delete Answer", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
  if res != DialogResult.Yes:
    return

  q.Answers.Remove(a)
  _renumber_answers(q)
  _ensure_selected_context(vm)
  
  vm.Debug("Deleted answer '{0}' from question '{1}'".format(a.AnswerText, q.QuestionText))
  log_line("Deleted answer '{0}' from question '{1}'".format(a.AnswerText, q.QuestionText))  
  return

# --- Logging and VM Dumping Utilities ---
def _ensure_dir(path):
  if not Directory.Exists(path):
    Directory.CreateDirectory(path)

def _safe_username():
  try:
    return Environment.UserName
  except:
    return "unknown_user"

def _safe_machine():
  try:
    return Environment.MachineName
  except:
    return "unknown_machine"

def get_Database_Name():
  # returns current database name (Live = 'Partner', Dev = 'PartnerDev' and Training = 'PartnerTraining')
  activeDB = runSQL(codeToRun="SELECT DB_NAME()", returnType="String")
  return activeDB

def _get_Snapshot_dir():
  # structure: \\ParnerUNC\MRA_Snapshots\YYYY-MM-DD\
  ts = DateTime.Now.ToString("yyyy-MM-dd")
  SNAPSHOT_ROOT = r"\\tw-p4wapp01\{0}\Managing Partner\Forms\NMRA and FileReviews\MRA_Snapshots".format(get_Database_Name())
  snap_dir = Path.Combine(SNAPSHOT_ROOT, ts)

  try:
    _ensure_dir(snap_dir)
    return snap_dir
  except:
    # network share down / permissions issue -> fallback local
    fb = Path.Combine(_local_fallback_root(), "Snapshots", ts)
    _ensure_dir(fb)
    return fb

def _local_fallback_root():
  # e.g. C:\Users\you\AppData\Local\Temp\MRA_Log
  return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Temp", "MRA_Log")

def _get_log_dir():
  # structure: \\ParnerUNC\MRA_Log\YYYY-MM-DD\
  today = DateTime.Now.ToString("yyyy-MM-dd")
  LOG_ROOT = r"\\tw-p4wapp01\{0}\Managing Partner\Forms\NMRA and FileReviews\MRA_Log".format(get_Database_Name())
  net_dir = Path.Combine(LOG_ROOT, today)

  try:
    _ensure_dir(net_dir)
    return net_dir
  except:
    # network share down / permissions issue -> fallback local
    fb = Path.Combine(_local_fallback_root(), today)
    _ensure_dir(fb)
    return fb

def log_line(message, template_id=None):
  # Appends a line to a daily per-user log file.
  log_dir = _get_log_dir()

  user = _safe_username()
  machine = _safe_machine()
  file_name = "{0}_{1}.log".format(user, machine)  # per user+machine
  path = Path.Combine(log_dir, file_name)

  ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")
  tid = "" if template_id is None else " TID={0}".format(template_id)
  line = "{0}{1} {2}\r\n".format(ts, tid, str(message))

  # append with basic retry
  try:
    File.AppendAllText(path, line)
  except Exception as ex:
    # last-resort: try local fallback
    fb_dir = Path.Combine(_local_fallback_root(), DateTime.Now.ToString("yyyy-MM-dd"))
    _ensure_dir(fb_dir)
    fb_path = Path.Combine(fb_dir, file_name)
    File.AppendAllText(fb_path, line)

  return path  # return path for convenience


def btn_EditMRA_CopyDebugLog_Click(sender, e):
  vm = _tikitSender.DataContext
  if vm is None:
    return

  # Copy debug log to clipboard
  log_text = "\r\n".join([str(x) for x in vm.DebugLines])
  Clipboard.SetText(log_text)
  MessageBox.Show("Debug log copied to clipboard.", "Debug Log Copied", MessageBoxButtons.OK, MessageBoxIcon.Information)
  return


def btn_DumpVM_Click(sender, e):
  vm = _tikitSender.DataContext
  if vm is None:
    return
  path = dump_vm_to_file(vm)

  # ask user if they want to open the file
  result = MessageBox.Show("The MRA structure has been exported to:\n{0}\n\nDo you want to open the file now?".format(path), "VM Dump Created", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
  if result == DialogResult.Yes:
    try:
      Process.Start("notepad.exe", path)
    except Exception as ex:
      MessageBox.Show("Could not open the file automatically:\n{0}".format(str(ex)), "Error Opening File", MessageBoxButtons.OK, MessageBoxIcon.Error)
  return

def dump_template_vm_text(vm):
  lines = []
  lines.append("TemplateID={0}".format(vm.TemplateID))
  lines.append("Groups={0}".format(vm.Groups.Count))

  for g in vm.Groups:
    lines.append("")
    lines.append("GROUP: '{0}'\nQuestions={1}".format(g.GroupName, g.Questions.Count))
    for q in g.Questions:
      lines.append("  Q#{0}  QID={1}  '{2}'".format(q.QuestionDisplayOrder, q.QuestionID, q.QuestionText))
      for a in q.Answers:
        lines.append("    A#{0}  AID={1}  Score={2}  Text='{3}'  EmailComment='{4}'".format(
                  a.AnswerDisplayOrder, a.AnswerID, a.Score, a.AnswerText, a.EmailComment
              ))
  return "\r\n".join(lines)


def dump_vm_to_file(vm, prefix="USER_TRIGGERED"):
  # for UI button we'll prefix 'USER_TRIGGERED' to distinguish from automated dumps in logs (which will prefix with 'AUTO_BEFORE' or 'AUTO_AFTER')
  log_dir = _get_Snapshot_dir()
  ts = DateTime.Now.ToString("yyyyMMdd_HHmmssfff")
  user = _safe_username()
  machine = _safe_machine()

  fname = "TID{0}_{1}_{2}_{3}.txt".format(vm.TemplateID, prefix, user, machine)
  # Add timestamp to avoid overwrites
  fname = "{0}_{1}.txt".format(fname.replace(".txt", ""), ts)

  path = Path.Combine(log_dir, fname)

  content = dump_template_vm_text(vm)
  File.WriteAllText(path, content)

  log_line("VM dump written: {0}".format(path), vm.TemplateID)
  return path


def btn_EditMRA_SaveToDB_Click(sender, e):
  vm = _tikitSender.DataContext
  if vm is None:
    return

  # Validate VM first (no DB changes yet)
  ok, msg = validate_unique_question_text(vm)
  if not ok:
    MessageBox.Show(msg, "Cannot Save Template", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    vm.Debug("SAVE BLOCKED: " + msg)
    log_line("SAVE BLOCKED: " + msg, to_int(vm.TemplateID))
    return

  # Confirm save action
  res = MessageBox.Show(
      "Are you sure you want to save the current template structure to the database?\n\n"
      "This will overwrite the existing structure for this template.",
      "Confirm Save to Database",
      MessageBoxButtons.YesNo,
      MessageBoxIcon.Question
  )
  if res != DialogResult.Yes:
    return

  # Snapshot "after" (attempted state)
  dump_vm_to_file(vm, prefix="AUTO_AFTER")

  try:
    save_template_to_db(vm)  # <-- no UI inside, just raises on error
    MessageBox.Show("Template saved successfully.", "Save Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)
    vm.Debug("TemplateID {0} saved to database successfully.".format(to_int(vm.TemplateID)))
    log_line("TemplateID {0} saved to database successfully.".format(to_int(vm.TemplateID)), to_int(vm.TemplateID))
  except Exception as ex:
    MessageBox.Show("An error occurred while saving the template:\n\n{0}".format(str(ex)), "Save Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
    vm.Debug("Error saving TemplateID {0}: {1}".format(to_int(vm.TemplateID), str(ex)))
    log_line("Error saving TemplateID {0}: {1}".format(to_int(vm.TemplateID), str(ex)), to_int(vm.TemplateID))

  return


def save_template_to_db(vm):
  # This function will save the template structure from the ViewModel (vm) back to the database.
  if vm is None:
    return

  tid = to_int(vm.TemplateID)
 
  # 1a) Resolve Questions/Answers (add any new ones,  handle duplicates as needed)
  # Note: this logs results to Debug window AND the log_file
  for g in vm.Groups:
    for q in g.Questions:
      #resolve_question_id(vm, q)
      get_or_create_question_id(vm, q)
      for a in q.Answers:
        #resolve_answer_id(vm, a)
        get_or_create_answer_id(vm, a)

  # 1b) Flatten rows (output list of dicts with TemplateID, QuestionID, AnswerID, Score, QuestionGroup, QuestionOrder, 
  # AnswerOrder for each answer (or question if no answers)) - NB: outputs status to log/debug
  rows = flatten_template_rows(vm)

    # 2) Build 'apply' statements for the 'write' phase
  sql_batch = []
  sql_batch.append("DELETE FROM Usr_MRAv2_Templates WHERE TemplateID = {0};".format(tid))

  if rows:
    values_sql = []
    for r in rows:
      values_sql.append(
        "({TemplateID}, {QuestionID}, {AnswerID}, {Score}, '{QuestionGroup}', {QuestionOrder}, {AnswerOrder})".format(
            TemplateID=to_int(r["TemplateID"]),
            QuestionID=to_int(r["QuestionID"]),
            AnswerID=to_int(r["AnswerID"]),
            Score=to_int(r["Score"]),
            QuestionGroup=sql_escape(r["QuestionGroup"]),
            QuestionOrder=to_int(r["QuestionOrder"]),
            AnswerOrder=to_int(r["AnswerOrder"]),
          )
        )

    sql_batch.append(
        "INSERT INTO Usr_MRAv2_Templates "
        "(TemplateID, QuestionID, AnswerID, Score, QuestionGroup, QuestionOrder, AnswerOrder) "
        "VALUES {0};".format(", ".join(values_sql))
    )
    vm.Debug("Prepared INSERT for {0} rows".format(len(rows)))
    log_line("Prepared INSERT for {0} rows.".format(len(rows)))
      
  run_apply_phase(vm, sql_batch)

  vm.Debug("SAVE Complete for TemplateID: {0}".format(tid))
  log_line("SAVE Complete for TemplateID: {0}".format(tid))
  return


def run_apply_phase(vm, sql_batch):
  tid = to_int(vm.TemplateID)
  if not sql_batch:
    return

  # Preferred: execute as ONE batch, so it is at least same session *if* your provider keeps it
  # (still not a real transaction, but reduces partial apply risk)
  batch_sql = ";\r\n".join(sql_batch) + ";"

  vm.Debug("APPLY batch for TID={0}:\n{1}".format(tid, batch_sql))
  log_line("APPLY batch for TID={0}".format(tid), tid)

  try:
    # Try resolver first (often better at multi-statement batches)
    runSQL(batch_sql, useAltResolver=False, returnType="String")
    vm.Debug("APPLY batch executed successfully for TID={0}".format(tid))
    log_line("APPLY batch executed successfully for TID={0}".format(tid), tid)
    return
  except Exception as ex:
    vm.Debug("APPLY batch failed, falling back to per-statement. Error: " + str(ex))
    log_line("APPLY batch failed, falling back to per-statement. Error: " + str(ex), tid)

  # Fallback: execute each statement separately
  for i, stmt in enumerate(sql_batch):
    vm.Debug("APPLY stmt {0}/{1}: {2}".format(i+1, len(sql_batch), stmt))
    runSQL(stmt, useAltResolver=False, returnType="String")
  return


def flatten_template_rows(vm):
  """
  Returns a list of dicts for use with 'INSERT INTO Usr_MRAv2_Templates (...) VALUES ...' with one dict per answer (or question if no answers), 
  containing all necessary info to reconstruct the template structure in the database, including display orders:
    {TemplateID, QuestionID, AnswerID, Score, QuestionGroup, QuestionOrder, AnswerOrder}
  """
  rows = []
  tid = to_int(vm.TemplateID)

  for g in vm.Groups:
    group_name = g.GroupName or ""

    q_order = 1
    for q in g.Questions:
      if to_int(q.QuestionDisplayOrder) != q_order:
        q.QuestionDisplayOrder = q_order

      a_order = 1
      for a in q.Answers:
        if to_int(a.AnswerDisplayOrder) != a_order:
          a.AnswerDisplayOrder = a_order

        rows.append({
          "TemplateID": tid,
          "QuestionID": to_int(q.QuestionID),
          "AnswerID": to_int(a.AnswerID),
          "Score": to_int(a.Score, 0),
          "QuestionGroup": group_name,
          "QuestionOrder": to_int(q.QuestionDisplayOrder),
          "AnswerOrder": to_int(a.AnswerDisplayOrder),
        })
        a_order += 1

      q_order += 1

  vm.Debug("Flattened template structure into {0} rows for database insertion.".format(len(rows)))
  log_line("Flattened template structure into {0} rows for database insertion.".format(len(rows)), tid)
  return rows


def norm_text(s):
  # basic normalisation; keep it conservative
  s = "" if s is None else str(s)
  s = s.strip()
  # optional: collapse internal whitespace
  # s = " ".join(s.split())
  return s

def get_or_create_question_id(vm, q):

  txt = norm_text(q.QuestionText)
  if txt == "":
    raise Exception("Question text cannot be blank.")

  # 1) lookup by text
  existing = db_scalar(
    "SELECT TOP 1 QuestionID FROM Usr_MRAv2_Question WHERE QuestionText = '{0}'".format(sql_escape(txt))
  )
  eid = to_int(existing, 0)
  if eid > 0:
    q.QuestionID = eid
    return eid

  # 2) insert
  new_id = db_scalar(
    "INSERT INTO Usr_MRAv2_Question (QuestionText, QuestionID) OUTPUT INSERTED.QuestionID SELECT '{0}', MAX(QuestionID) + 1 FROM Usr_MRAv2_Question".format(sql_escape(txt))
  )
  new_id = to_int(new_id, 0)
  if new_id <= 0:
    raise Exception("Failed to insert Question.")
  q.QuestionID = new_id
  return new_id

def get_or_create_answer_id(vm, a):
  txt = norm_text(a.AnswerText)
  email = norm_text(a.EmailComment)

  if txt == "":
    raise Exception("Answer text cannot be blank.")

  # 1) lookup by text+email (because you store EmailComment in the Answer table)
  existing = db_scalar(
    "SELECT TOP 1 AnswerID FROM Usr_MRAv2_Answer "
    "WHERE AnswerText = '{0}' AND ISNULL(EmailComment,'') = '{1}'".format(
      sql_escape(txt),
      sql_escape(email)
    )
  )
  eid = to_int(existing, 0)
  if eid > 0:
    a.AnswerID = eid
    return eid

  # 2) insert
  new_id = db_scalar(
    "INSERT INTO Usr_MRAv2_Answer (AnswerText, EmailComment, AnswerID) OUTPUT INSERTED.AnswerID SELECT '{0}', '{1}', MAX(AnswerID) + 1 FROM Usr_MRAv2_Answer".format(sql_escape(txt), sql_escape(email))
  )
  new_id = to_int(new_id, 0)
  if new_id <= 0:
    raise Exception("Failed to insert Answer.")
  a.AnswerID = new_id
  return new_id


def resolve_question_id(vm, q):
  """
  Ensures q.QuestionID is correct for saving:
    - If new (negative): insert, set new ID
    - If existing and text differs:
        * if used elsewhere -> prompt update globally vs clone
        * else update directly
  Returns final QuestionID.
  """
  qid = to_int(q.QuestionID)
  new_text = q.QuestionText or ""

  vm.Debug("Resolving QuestionID for question '{0}' (current ID: {1})".format(new_text, qid))
  log_line("Resolving QuestionID for question '{0}' (current ID: {1})".format(new_text, qid))

  # compare current DB text if we already have ID to see if text was changed
  if qid > 0:
    db_text = db_scalar(
        "SELECT QuestionText FROM Usr_MRAv2_Question WHERE QuestionID = {0};".format(qid)
    )
    db_text = "" if db_text is None else str(db_text)

    if db_text == new_text:
      return qid  # no change

    # Check usage in other templates (excluding current template)
    used_elsewhere = db_scalar(
        "SELECT COUNT(*) FROM Usr_MRAv2_Templates "
        "WHERE QuestionID = {0} AND TemplateID <> {1};".format(qid, to_int(vm.TemplateID))
    )
    used_elsewhere = 0 if used_elsewhere is None else to_int(used_elsewhere)

    if used_elsewhere <= 0:
      # safe: update in place
      db_nonquery(
          "UPDATE Usr_MRAv2_Question SET QuestionText = '{0}' WHERE QuestionID = {1};".format(
              sql_escape(new_text), qid
          )
      )
      vm.Debug("Updated question text for QuestionID={0} since not used in other templates.".format(qid))
      log_line("Updated question text for QuestionID={0} since not used in other templates.".format(qid))
      return qid

    # Used elsewhere: prompt
    res = MessageBox.Show(
        "Question text has been changed, and this QuestionID is used in {0} other template(s).\n\n"
        "YES  = Update ALL templates (update Usr_MRAv2_Question)\n"
        "NO   = Create a NEW question (only this template will use it)\n"
        "CANCEL = Abort save".format(used_elsewhere),
        "Question text changed",
        MessageBoxButtons.YesNoCancel,
        MessageBoxIcon.Warning
    )

    if res == DialogResult.Cancel:
        vm.Debug("Save cancelled by user due to question text change for QuestionID={0}".format(qid))
        log_line("Save cancelled by user due to question text change for QuestionID={0}".format(qid))
        raise Exception("Save cancelled by user (question change).")

    if res == DialogResult.Yes:
        db_nonquery(
            "UPDATE Usr_MRAv2_Question SET QuestionText = '{0}' WHERE QuestionID = {1};".format(
                sql_escape(new_text), qid
            )
        )
        vm.Debug("Updated question text for QuestionID={0}. User requested update 'all' other occurrences.".format(qid))
        log_line("Updated question text for QuestionID={0}. User requested update 'all' other occurrences.".format(qid))
        return qid

  # Not previously used / user answered 'no' to 'updating' existing question text, so clone:
  insert_sql = (
      "INSERT INTO Usr_MRAv2_Question (QuestionText, QuestionID) "
      "OUTPUT INSERTED.QuestionID "
      "SELECT '{0}', MAX(QuestionID) + 1 FROM Usr_MRAv2_Question;"
  ).format(sql_escape(new_text))
  new_id = db_scalar(insert_sql)
  q.QuestionID = to_int(new_id)
  vm.Debug("Inserted new question '{0}' with QuestionID={1} since user requested to clone due to text change.".format(new_text, new_id))
  log_line("Inserted new question '{0}' with QuestionID={1} since user requested to clone due to text change.".format(new_text, new_id))
  return to_int(new_id)


def resolve_answer_id(vm, a):
  """
  Ensures a.AnswerID is correct for saving:
    - If new (negative): insert, set new ID
    - If existing and text/email differs:
        * if used elsewhere -> prompt update globally vs clone
        * else update directly
  Returns final AnswerID.
  """
  aid = to_int(a.AnswerID)
  new_text = a.AnswerText or ""
  new_email = a.EmailComment or ""

  vm.Debug("Resolving AnswerID for answer '{0}' (current ID: {1})".format(new_text, aid))
  log_line("Resolving AnswerID for answer '{0}' (current ID: {1})".format(new_text, aid))

  # compare current DB text/email
  db_row = None
  _tikitDbAccess.Open(
      "SELECT AnswerText, EmailComment FROM Usr_MRAv2_Answer WHERE AnswerID = {0};".format(aid)
  )
  dr = _tikitDbAccess._dr
  if dr is not None and dr.HasRows and dr.Read():
    t = "" if dr.IsDBNull(0) else dr.GetString(0)
    e = "" if dr.IsDBNull(1) else dr.GetString(1)
    db_row = (t, e)
    dr.Close()
  _tikitDbAccess.Close()

  if db_row is None:
    # If somehow missing in DB, treat like new
    insert_sql = (
        "INSERT INTO Usr_MRAv2_Answer (AnswerText, EmailComment, AnswerID) "
        "OUTPUT INSERTED.AnswerID "
        "SELECT '{0}', '{1}', MAX(AnswerID) + 1 FROM Usr_MRAv2_Answer;"
    ).format(sql_escape(new_text), sql_escape(new_email))
    new_id = db_scalar(insert_sql)
    a.AnswerID = to_int(new_id)
    vm.Debug("Inserted new answer '{0}' with AnswerID={1}. No match to existing in DB.".format(new_text, new_id))
    log_line("Inserted new answer '{0}' with AnswerID={1}. No match to existing in DB.".format(new_text, new_id))   
    return to_int(new_id)


  db_text, db_email = db_row
  if db_text == new_text and (db_email or "") == (new_email or ""):
    vm.Debug("No change to answer '{0}' with AnswerID={1}. Matches existing in DB.".format(new_text, aid))
    log_line("No change to answer '{0}' with AnswerID={1}. Matches existing in DB.".format(new_text, aid))   
    return aid  # no change

  used_elsewhere = db_scalar(
      "SELECT COUNT(*) FROM Usr_MRAv2_Templates "
      "WHERE AnswerID = {0} AND TemplateID <> {1};".format(aid, vm.TemplateID)
  )
  used_elsewhere = 0 if used_elsewhere is None else to_int(used_elsewhere)

  if used_elsewhere <= 0:
    # not used elsewhere: safe to update in place
    db_nonquery(
        "UPDATE Usr_MRAv2_Answer SET AnswerText = '{0}', EmailComment = '{1}' WHERE AnswerID = {2};".format(
            sql_escape(new_text), sql_escape(new_email), aid
        )
    )
    vm.Debug("Updated answer text/email for AnswerID={0} since not used in other templates.".format(aid))
    log_line("Updated answer text/email for AnswerID={0} since not used in other templates.".format(aid))
    return aid

  # used elsewhere prompt:
  res = MessageBox.Show(
      "Answer text/comment has been changed, and this AnswerID is used in {0} other template(s).\n\n"
      "YES  = Update ALL templates (update Usr_MRAv2_Answer)\n"
      "NO   = Create a NEW answer (only this template will use it)\n"
      "CANCEL = Abort save".format(used_elsewhere),
      "Answer changed",
      MessageBoxButtons.YesNoCancel,
      MessageBoxIcon.Warning
  )

  if res == DialogResult.Cancel:
    vm.Debug("Save cancelled by user due to answer text/email change for AnswerID={0}".format(aid))
    log_line("Save cancelled by user due to answer text/email change for AnswerID={0}".format(aid))
    raise Exception("Save cancelled by user (answer change).")

  if res == DialogResult.Yes:
    db_nonquery(
        "UPDATE Usr_MRAv2_Answer SET AnswerText = '{0}', EmailComment = '{1}' WHERE AnswerID = {2};".format(
            sql_escape(new_text), sql_escape(new_email), aid
        )
    )
    vm.Debug("Updated answer text/email for AnswerID={0}. User requested update 'all' other occurrences.".format(aid))
    log_line("Updated answer text/email for AnswerID={0}. User requested update 'all' other occurrences.".format(aid))
    return aid

  # NO: clone
  insert_sql = (
      "INSERT INTO Usr_MRAv2_Answer (AnswerText, EmailComment, AnswerID) "
      "OUTPUT INSERTED.AnswerID "
      "SELECT '{0}', '{1}', MAX(AnswerID) + 1 FROM Usr_MRAv2_Answer;"
    ).format(sql_escape(new_text), sql_escape(new_email))
  new_id = db_scalar(insert_sql)
  a.AnswerID = to_int(new_id)
  vm.Debug("Inserted new answer '{0}' with AnswerID={1} since user requested to clone due to text/email change.".format(new_text, new_id))
  log_line("Inserted new answer '{0}' with AnswerID={1} since user requested to clone due to text/email change.".format(new_text, new_id))
  return to_int(new_id)


## Helper function for SQL
def db_scalar(sql):
  # Returns first column of first row or None.
  # In theory, we could just use TikitResolver for this, but to keep it consistent with our db access patterns and ensure proper opening/closing, we'll do it manually here.
  val = None
  _tikitDbAccess.Open(sql)
  dr = _tikitDbAccess._dr
  if dr is not None and dr.HasRows:
    if dr.Read():
      if not dr.IsDBNull(0):
        val = dr.GetValue(0)
    dr.Close()
  _tikitDbAccess.Close()
  return val

def db_nonquery(sql):
  # Executes non-query SQL.
  _tikitDbAccess.Open(sql)
  # Partner/Tikit often executes on Open for non-select; if you have a dedicated method use that instead.
  #! We do have 'runSQL'
  # Ensure reader closed if any
  dr = _tikitDbAccess._dr
  if dr is not None:
    try:
      dr.Close()
    except:
      pass
  _tikitDbAccess.Close()


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

def to_str(x, default=""):
  if is_dbnull(x):
    return default
  try:
    return str(x)
  except:
    return default


def validate_unique_question_text(vm):
  # Returns (ok, message)
  # Rule: within a template, QuestionText must be unique (case/trim normalised)
  seen = {}   # norm_text -> first QuestionVM
  dups = []   # list of tuples: (text, first_q, dup_q)

  for g in vm.Groups:
    for q in g.Questions:
      t = norm_text(q.QuestionText)
      if t == "":
        return (False, "One or more questions have blank text. Please complete or delete blank questions before saving.")
      if t in seen:
        dups.append((t, seen[t], q))
      else:
        seen[t] = q

  if not dups:
    return (True, "")

  # Build a helpful error message
  # Include group + display order to help user find them
  lines = []
  lines.append("Duplicate QuestionText found in this template. Question text must be unique before saving.\n")
  for (t, q1, q2) in dups[:20]:
    g1 = q1.ParentGroup.GroupName if getattr(q1, "ParentGroup", None) is not None else "?"
    g2 = q2.ParentGroup.GroupName if getattr(q2, "ParentGroup", None) is not None else "?"
    lines.append("• '{0}'".format(t))
    lines.append("    - First:  Group='{0}', Order={1}, QID={2}".format(g1, to_int(q1.QuestionDisplayOrder), to_int(q1.QuestionID)))
    lines.append("    - Again:  Group='{0}', Order={1}, QID={2}".format(g2, to_int(q2.QuestionDisplayOrder), to_int(q2.QuestionID)))
  if len(dups) > 20:
    lines.append("\n(Showing first 20 duplicates; there are {0} total.)".format(len(dups)))

  return (False, "\n".join(lines))

######################################################################################################################

## P R E V I E W   M R A   T E M P L A T E   ##

def btn_MRATemplate_Preview_Click(s, event):
  # This function will load the 'Preview' tab (made to look like 'matter-level' XAML) for the selected item

  # if nothing selected, alert user and quit
  if lbl_MRATemplate_ID.Content == '0' or dg_MRA_Templates.SelectedIndex == UNSELECTED:
    MessageBox.Show("Nothing selected to Preview!", "Error: Preview selected Matter Risk Assessment...")
    return
  
  # put details into header area of 'Preview' tab
  tb_MRAPreview_ID.Text = str(lbl_MRATemplate_ID.Content)
  tb_MRAPreview_Name.Text = str(tb_MRATemplate_Name.Text)
  tb_MRAPreview_Score.Text = '0'          #str(tb_ScoreLowTo.Text)
  tb_MRAPreview_RiskCategory.Text = '-'   # "Low Risk" if int(tb_ScoreLowTo.Text) == 0 else "Medium Risk" if int(tb_ScoreMedFrom.Text) > 0 and int(tb_ScoreMedFrom.Text) <= int(tb_ScoreHighFrom.Text) else "High Risk"
  tb_MRAPreview_RiskCategoryID.Text = "1" if tb_MRAPreview_RiskCategory.Text == "Low Risk" else "2" if tb_MRAPreview_RiskCategory.Text == "Medium Risk" else "3"
  tb_MRAPreview_ScoreTriggerMedium.Text = str(tb_ScoreMedFrom.Text)
  tb_MRAPreview_ScoreTriggerHigh.Text = str(tb_ScoreHighFrom.Text)

  # get AnswerList in memory for this template
  MRAPreview_load_Answers_toMemory()

  # load DataGrid with Questions
  MRAPreview_load_Questions_DataGrid()

  # finally, show this Preview tab and hide 'Overview' tab
  ti_MRA_Overview.Visibility = Visibility.Collapsed
  ti_MRA_Preview.Visibility = Visibility.Visible
  ti_MRA_Preview.IsSelected = True
  #MessageBox.Show("Preview selected MRA click", "Preview selected Matter Risk Assessment...")
  return
  

def btn_MRAPreview_BackToOverview_Click(sender, e):
  # Just switch back to overview tab; VM is already updated since we're binding directly to the edit fields
  
  ti_MRA_Preview.Visibility = Visibility.Collapsed
  ti_MRA_Overview.Visibility = Visibility.Visible
  ti_MRA_Overview.IsSelected = True
  return

# Concept with Preview MRA is to load MRA Template into memory only - and we get answer data from DataGrid data
# difference from v1 is that we have different columns in the datagrid, and got rid of redundant ones.

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
class QuestionItem(NotifyBase):
  def __init__(self, group_name, order_no, qid, qtext, answers):
    NotifyBase.__init__(self)
    self.QuestionGroup = group_name or ""
    self.QuestionOrder = int(order_no) if order_no is not None else 0
    self.QuestionID = int(qid) if qid is not None else None
    self.QuestionText = qtext or ""
    self.AvailableAnswers = answers or []

    self._SelectedAnswer = None  # <-- bind target

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
  view = CollectionViewSource.GetDefaultView(dg_MRAPreview.ItemsSource)
  if view is None:
    return None
  return view.CurrentItem


def MRAPreview_load_Answers_toMemory():
  global MRA_PREVIEW_ANSWERS_BY_QID
  MRA_PREVIEW_ANSWERS_BY_QID = {}

  mySQL = """SELECT T.QuestionID, Ans.AnswerID, Ans.AnswerText, Ans.EmailComment, T.Score
             FROM Usr_MRAv2_Templates T
                JOIN Usr_MRAv2_Answer Ans ON T.AnswerID = Ans.AnswerID
             WHERE T.TemplateID = {0}
             ORDER BY T.QuestionID, T.AnswerOrder;""".format(tb_MRAPreview_ID.Text)

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
      MRA_PREVIEW_ANSWERS_BY_QID.setdefault(qid, []).append(item)

    dr.Close()
  _tikitDbAccess.Close()
  return

def MRAPreview_load_Questions_DataGrid():
  # This function will populate the Matter Risk Assessment Preview datagrid
  #MessageBox.Show("Start - getting group ID", "Refreshing list (datagrid of questions)")

  global MRA_PREVIEW_QUESTIONS_LIST
  # firstly, wipe list in case we're reloading
  MRA_PREVIEW_QUESTIONS_LIST = []

  #MessageBox.Show("Genating SQL...", "Refreshing list (datagrid of questions)")
  mySQL = """SELECT MRAT.QuestionGroup, MRAT.QuestionOrder, MRAT.QuestionID, MRAQ.QuestionText
             FROM Usr_MRAv2_Templates MRAT
                LEFT JOIN Usr_MRAv2_Question MRAQ ON MRAT.QuestionID = MRAQ.QuestionID
             WHERE MRAT.TemplateID = {0}
             GROUP BY MRAT.QuestionGroup, MRAT.QuestionOrder, MRAT.QuestionID, MRAQ.QuestionText
             ORDER BY MRAT.QuestionGroup, MRAT.QuestionOrder;""".format(tb_MRAPreview_ID.Text)
  
  #MessageBox.Show("SQL: " + str(mySQL) + "\n\nRefreshing list (datagrid of questions)", "Debug: Populating List of Questions (Preview MRA)")

  _tikitDbAccess.Open(mySQL)
  dr = _tikitDbAccess._dr
  if dr is not None and dr.HasRows:
    while dr.Read():
      group_name = "" if dr.IsDBNull(0) else dr.GetValue(0)
      order_no = to_int(dr.GetValue(1))
      qid = to_int(dr.GetValue(2))
      qtext = "" if dr.IsDBNull(3) else dr.GetString(3)

      answers = MRA_PREVIEW_ANSWERS_BY_QID.get(qid, [])
      MRA_PREVIEW_QUESTIONS_LIST.append(QuestionItem(group_name, order_no, qid, qtext, answers))

    dr.Close()
  _tikitDbAccess.Close()
  
  # and now we have a list of question items in memory for this template, with their available answers;
  # we can bind this to the datagrid and it should show the questions grouped by 'QuestionGroup' with a
  # combo box of available answers for each question (bound to the 'SelectedAnswerID' property of the
  # QuestionItem, which will allow us to easily get the selected answer and its score/email comment (when
  # user selects an answer in the preview)
  # create observable collection for WPF and bind to datagrid; this should show the questions grouped by 'QuestionGroup' with a combo box of available answers for each question (bound to the 'SelectedAnswerID' property of the QuestionItem, which will allow us to easily get the selected answer and its score/email comment when user selects an answer in the preview)
  view = ListCollectionView(MRA_PREVIEW_QUESTIONS_LIST)
  view.GroupDescriptions.Add(PropertyGroupDescription("QuestionGroup"))
  dg_MRAPreview.ItemsSource = view

  has_items = (len(MRA_PREVIEW_QUESTIONS_LIST) > 0)
  grid_Preview_MRA.Visibility = Visibility.Visible if has_items else Visibility.Collapsed
  tb_NoMRA_PreviewQs.Visibility = Visibility.Collapsed if has_items else Visibility.Visible

  if has_items:
    # defer until UI has built group containers
    dg_MRAPreview.Dispatcher.BeginInvoke(
                  DispatcherPriority.ContextIdle,
                  Action(_select_first_preview_row)
                  )

    #first = MRA_PREVIEW_QUESTIONS_LIST[0]   # <-- this is the first QUESTION row
    #dg_MRAPreview.SelectedItem = first      # <-- avoids group header selection
    #view.MoveCurrentTo(first)               # <-- keeps CurrentItem in sync with SelectedItem
    #dg_MRAPreview.ScrollIntoView(first)
  return


def _select_first_preview_row():
  if len(MRA_PREVIEW_QUESTIONS_LIST) <= 0:
    return

  first = MRA_PREVIEW_QUESTIONS_LIST[0]

  # force containers to exist
  try:
    dg_MRAPreview.UpdateLayout()
  except:
    pass

  # 1) Force a real selection change event (important with Grouping DataGrids), select nothing and then first row
  dg_MRAPreview.SelectedItem = None
  dg_MRAPreview.SelectedItem = first

  # 2) ensure CurrentItem is aligned
  view = _get_preview_view()
  if view is not None:
    try:
      view.MoveCurrentTo(first)
    except:
      pass

  # 3) Commit "current cell" so WPF treats it as a real row selection
  try:
    if dg_MRAPreview.Columns.Count > 0:
      dg_MRAPreview.CurrentCell = DataGridCellInfo(first, dg_MRAPreview.Columns[0])
  except:
    pass

  try:
    dg_MRAPreview.ScrollIntoView(first)
    dg_MRAPreview.Focus()
  except:
    pass

  _sync_combo_to_current_row()
  MRAPreview_RecalcTotalScore()


def MRAPreview_AdvanceToNextQuestion():
  # Uses the real list, so grouping doesn't break indices
  if len(MRA_PREVIEW_QUESTIONS_LIST) <= 0:
    return

  curr = _get_current_question()
  if curr is None:
    # if something went odd, just go to first
    _select_first_preview_row()
    return

  try:
    idx = MRA_PREVIEW_QUESTIONS_LIST.index(curr)
  except:
    # CurrentItem might be a group wrapper; fall back to SelectedItem
    try:
      idx = MRA_PREVIEW_QUESTIONS_LIST.index(dg_MRAPreview.SelectedItem)
    except:
      _select_first_preview_row()
      return

  next_idx = idx + 1
  if next_idx >= len(MRA_PREVIEW_QUESTIONS_LIST):
    next_idx = 0  # wrap to first; change this if you'd rather stop at end

  nxt = MRA_PREVIEW_QUESTIONS_LIST[next_idx]

  # Update selection + CurrentItem + scroll, then sync right panel combo
  dg_MRAPreview.SelectedItem = nxt

  view = _get_preview_view()
  if view is not None:
    try:
      view.MoveCurrentTo(nxt)
    except:
      pass

  try:
    dg_MRAPreview.ScrollIntoView(nxt)
  except:
    pass

  _sync_combo_to_current_row()
  return
  

def dg_MRAPreview_SelectionChanged(s, event):
  # defer slightly so CurrentItem is updated, especially with grouping
  try:
    dg_MRAPreview.Dispatcher.BeginInvoke(
      DispatcherPriority.Background,
      Action(_sync_combo_to_current_row)
    )
  except:
    _sync_combo_to_current_row()
  return


def _current_preview_row():
  # Prefer CurrentItem (works properly with grouping)
  view = CollectionViewSource.GetDefaultView(dg_MRAPreview.ItemsSource)
  if view is not None:
    try:
      return view.CurrentItem
    except:
      pass
  return dg_MRAPreview.SelectedItem

def cbo_MRAPreview_SelectedComboAnswer_SelectionChanged(s, event):
  global _preview_combo_syncing
  if _preview_combo_syncing:
    return  # ignore programmatic sync changes

  q = _get_current_question()
  if q is None or not hasattr(q, "SelectedAnswer"):
    return

  ans = cbo_MRAPreview_SelectedComboAnswer.SelectedItem

  # IMPORTANT: ignore the transient "None" that occurs when ItemsSource swaps
  # unless the user genuinely cleared it (rare in your UI)
  if ans is None:
    return

  q.SelectedAnswer = ans

  view = _get_preview_view()
  if view is not None:
    view.Refresh()

  MRAPreview_RecalcTotalScore()
  ##MessageBox.Show("Combo Selection Changed! Selected Question: " + str(row.QuestionText) + "\nSelected AnswerID: " + str(row.SelectedAnswerID) + "\nSelected Answer Text: " + str(row.SelectedAnswerText) + "\nSelected Answer Score: " + str(row.SelectedAnswerScore) + "\nSelected Answer Email Comment: " + str(row.SelectedAnswerEmailComment))
  return

def btn_MRAPreview_SaveAnswer_Click(s, event):
  q = _current_preview_row()
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
    auto = (chk_MRAPreview_AutoSelectNext.IsChecked == True)
  except:
    auto = False

  if auto:
    # Defer slightly so UI settles before we change selection
    try:
      dg_MRAPreview.Dispatcher.BeginInvoke(
        DispatcherPriority.Background,
        lambda: MRAPreview_AdvanceToNextQuestion()
      )
    except:
      MRAPreview_AdvanceToNextQuestion()

  return

## Helper functions to sync the ComboBox to datagrid (in terms of displaying value)
def _get_preview_view():
  return CollectionViewSource.GetDefaultView(dg_MRAPreview.ItemsSource)

def _get_current_question():
  view = _get_preview_view()
  if view is not None:
    try:
      return view.CurrentItem
    except:
      pass
  return dg_MRAPreview.SelectedItem

def _sync_combo_to_current_row():
  global _preview_combo_syncing
  q = _get_current_question()
  if q is None or not hasattr(q, "AvailableAnswers"):
    return

  _preview_combo_syncing = True
  try:
    # Make sure combo is showing this row’s answers
    try:
      cbo_MRAPreview_SelectedComboAnswer.ItemsSource = q.AvailableAnswers
    except:
      pass

    target = getattr(q, "SelectedAnswer", None)
    if target is None:
      cbo_MRAPreview_SelectedComboAnswer.SelectedItem = None
      return

    # Choose matching item by AnswerID
    tid = getattr(target, "AnswerID", None)
    found = None
    for a in q.AvailableAnswers:
      if getattr(a, "AnswerID", None) == tid:
        found = a
        break

    cbo_MRAPreview_SelectedComboAnswer.SelectedItem = found
  finally:
    _preview_combo_syncing = False


def MRAPreview_RecalcTotalScore():
  total = 0
  try:
    for q in MRA_PREVIEW_QUESTIONS_LIST:
      # SelectedAnswerScore is 0 when no answer selected
      try:
        total += int(getattr(q, "SelectedAnswerScore", 0) or 0)
      except:
        pass
  except:
    total = 0

  tb_MRAPreview_Score.Text = str(total)

  # now work out 'category' based on score and thresholds; we have two thresholds: MediumFrom and HighFrom; if score is below MediumFrom, it's Low Risk; if it's between MediumFrom and HighFrom, it's Medium Risk; if it's above HighFrom, it's High Risk
  tb_MRAPreview_ScoreTriggerMedium.Text
  tb_MRAPreview_ScoreTriggerHigh.Text
  if total < to_int(tb_MRAPreview_ScoreTriggerMedium.Text): 
    category = "Low"
    categoryNum = 1 
  elif total >= to_int(tb_MRAPreview_ScoreTriggerMedium.Text) and total < to_int(tb_MRAPreview_ScoreTriggerHigh.Text):
    category = "Medium" 
    categoryNum = 2
  elif total >= to_int(tb_MRAPreview_ScoreTriggerHigh.Text):
    category = "High" 
    categoryNum = 3
  else: 
    category = "-"
    categoryNum = 0 
  
  tb_MRAPreview_RiskCategory.Text = category
  tb_MRAPreview_RiskCategoryID.Text = str(categoryNum)
  return total


# SELECT Ans.AnswerID, Ans.AnswerText, Ans.EmailComment, T.QuestionID
# FROM Usr_MRAv2_Answer Ans 
#   LEFT OUTER JOIN Usr_MRAv2_Templates T ON Ans.AnswerID = T.AnswerID
# WHERE T.TemplateID = 1



##########################################################################

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
btn_Clipboard_Copy = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Clipboard_Copy')
btn_Clipboard_Copy.Click += btn_Clipboard_Copy_Click
Clipboard_Popup_Copy = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'Clipboard_Popup_Copy')
Clipboard_Popup_Copy.Closed += Clipboard_Popup_Copy_Closed

btn_CopyToClipboard_QuestionOnly = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_CopyToClipboard_QuestionOnly')
btn_CopyToClipboard_QuestionOnly.Click += btn_CopyToClipboard_QuestionOnly_Click
btn_CopyToClipboard_QuestionAndAnswers = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_CopyToClipboard_QuestionAndAnswers')
btn_CopyToClipboard_QuestionAndAnswers.Click += btn_CopyToClipboard_QuestionAndAnswers_Click
btn_CopyToClipboard_AnswersAll = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_CopyToClipboard_AnswersAll')
btn_CopyToClipboard_AnswersAll.Click += btn_CopyToClipboard_AnswersAll_Click
btn_CopyToClipboard_AnswerOnly = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_CopyToClipboard_AnswerOnly')
btn_CopyToClipboard_AnswerOnly.Click += btn_CopyToClipboard_AnswerOnly_Click
btn_CopyToClipboard_Clear = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_CopyToClipboard_Clear')
btn_CopyToClipboard_Clear.Click += btn_CopyToClipboard_Clear_Click

btn_Clipboard_Paste = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Clipboard_Paste')
btn_Clipboard_Paste.Click += btn_Clipboard_Paste_Click

tb_MRA_NoQuestionsText = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MRA_NoQuestionsText')
lbl_NoAnswers = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_NoAnswers')


## P R E V I E W   M A T T E R   R I S K   A S S E S S M E N T   - TAB ##
btn_MRAPreview_BackToOverview = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_MRAPreview_BackToOverview')
btn_MRAPreview_BackToOverview.Click += btn_MRAPreview_BackToOverview_Click

tb_MRAPreview_ID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MRAPreview_ID')
tb_MRAPreview_Name = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MRAPreview_Name')
tb_MRAPreview_Score = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MRAPreview_Score')
tb_MRAPreview_RiskCategory = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MRAPreview_RiskCategory')
tb_MRAPreview_RiskCategoryID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MRAPreview_RiskCategoryID')
tb_MRAPreview_ScoreTriggerMedium = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MRAPreview_ScoreTriggerMedium')
tb_MRAPreview_ScoreTriggerHigh = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MRAPreview_ScoreTriggerHigh')

tb_NoMRA_PreviewQs = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_NoMRA_PreviewQs')
grid_Preview_MRA = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'grid_Preview_MRA')

dg_MRAPreview = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_MRAPreview')
dg_MRAPreview.SelectionChanged += dg_MRAPreview_SelectionChanged

tb_MRAPreview_QuestionText = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MRAPreview_QuestionText')
cbo_MRAPreview_SelectedComboAnswer = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cbo_MRAPreview_SelectedComboAnswer')
cbo_MRAPreview_SelectedComboAnswer.SelectionChanged += cbo_MRAPreview_SelectedComboAnswer_SelectionChanged
btn_MRAPreview_SaveAnswer = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_MRAPreview_SaveAnswer')
btn_MRAPreview_SaveAnswer.Click += btn_MRAPreview_SaveAnswer_Click
chk_MRAPreview_AutoSelectNext = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'chk_MRAPreview_AutoSelectNext')
tb_MRAPreview_EC = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_MRAPreview_EC')

lbl_MRAPreview_DGID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRAPreview_DGID')
lbl_MRAPreview_CurrVal = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MRAPreview_CurrVal')



## New MRA Template Editor ViewModel ##
#templateEditorVM = None
tvTemplate = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tvTemplate')      # tree view listing Question Group > Question > Answers
tvTemplate.SelectedItemChanged += tvTemplate_SelectedItemChanged
btn_EditMRA_Group_Add = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_EditMRA_Group_Add')
btn_EditMRA_Group_Add.Click += btn_EditMRA_Group_Add_Click
btn_EditMRA_Question_Add = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_EditMRA_Question_Add')
btn_EditMRA_Question_Add.Click += btn_EditMRA_Question_Add_Click
btn_EditMRA_Answer_Add = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_EditMRA_Answer_Add')
btn_EditMRA_Answer_Add.Click += btn_EditMRA_Answer_Add_Click

btn_EditMRA_Group_MoveTop = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_EditMRA_Group_MoveTop')
btn_EditMRA_Group_MoveTop.Click += btn_EditMRA_Group_MoveTop_Click
btn_EditMRA_Group_MoveUp = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_EditMRA_Group_MoveUp')
btn_EditMRA_Group_MoveUp.Click += btn_EditMRA_Group_MoveUp_Click
btn_EditMRA_Group_MoveDown = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_EditMRA_Group_MoveDown')
btn_EditMRA_Group_MoveDown.Click += btn_EditMRA_Group_MoveDown_Click
btn_EditMRA_Group_MoveBottom = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_EditMRA_Group_MoveBottom')
btn_EditMRA_Group_MoveBottom.Click += btn_EditMRA_Group_MoveBottom_Click
btn_EditMRA_Group_DeleteSelected = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_EditMRA_Group_DeleteSelected')
btn_EditMRA_Group_DeleteSelected.Click += btn_EditMRA_Group_DeleteSelected_Click

btn_EditMRA_Question_MoveTop = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_EditMRA_Question_MoveTop')
btn_EditMRA_Question_MoveTop.Click += btn_EditMRA_Question_MoveTop_Click
btn_EditMRA_Question_MoveUp = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_EditMRA_Question_MoveUp')
btn_EditMRA_Question_MoveUp.Click += btn_EditMRA_Question_MoveUp_Click
btn_EditMRA_Question_MoveDown = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_EditMRA_Question_MoveDown')
btn_EditMRA_Question_MoveDown.Click += btn_EditMRA_Question_MoveDown_Click
btn_EditMRA_Question_MoveBottom = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_EditMRA_Question_MoveBottom')
btn_EditMRA_Question_MoveBottom.Click += btn_EditMRA_Question_MoveBottom_Click
btn_EditMRA_Question_DeleteSelected = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_EditMRA_Question_DeleteSelected')
btn_EditMRA_Question_DeleteSelected.Click += btn_EditMRA_Question_DeleteSelected_Click

btn_EditMRA_Answer_MoveTop = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_EditMRA_Answer_MoveTop')
btn_EditMRA_Answer_MoveTop.Click += btn_EditMRA_Answer_MoveTop_Click
btn_EditMRA_Answer_MoveUp = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_EditMRA_Answer_MoveUp')
btn_EditMRA_Answer_MoveUp.Click += btn_EditMRA_Answer_MoveUp_Click
btn_EditMRA_Answer_MoveDown = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_EditMRA_Answer_MoveDown')
btn_EditMRA_Answer_MoveDown.Click += btn_EditMRA_Answer_MoveDown_Click
btn_EditMRA_Answer_MoveBottom = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_EditMRA_Answer_MoveBottom')
btn_EditMRA_Answer_MoveBottom.Click += btn_EditMRA_Answer_MoveBottom_Click
btn_EditMRA_Answer_DeleteSelected = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_EditMRA_Answer_DeleteSelected')
btn_EditMRA_Answer_DeleteSelected.Click += btn_EditMRA_Answer_DeleteSelected_Click

btn_EditMRA_CopyDebugLog = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_EditMRA_CopyDebugLog')
btn_EditMRA_CopyDebugLog.Click += btn_EditMRA_CopyDebugLog_Click
btn_DumpVM = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_DumpVM')
btn_DumpVM.Click += btn_DumpVM_Click
btn_EditMRA_SaveToDB = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_EditMRA_SaveToDB')
btn_EditMRA_SaveToDB.Click += btn_EditMRA_SaveToDB_Click

# Define Actions and on load events
myOnLoadEvent(_tikitSender, 'onLoad')

]]>
    </Loaded>
  </RiskPracticeV2>

</tfb>